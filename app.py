import os, zipfile, io, shutil, re
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, flash
import mimetypes
from aip import AipOcr
from pdf2image import convert_from_path
from flask_sqlalchemy import SQLAlchemy
from config import BAIDU_CONFIG, POPPLER_PATH

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///invoices_pro.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = os.environ.get('FLASK_SECRET', 'devsecret')
db = SQLAlchemy(app)

# --- 关键：文件名清理函数，防止星号等非法字符报错 ---
def clean_path_name(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name)

class Invoice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    inv_num = db.Column(db.String(50))
    inv_code = db.Column(db.String(50))
    date = db.Column(db.String(20))
    seller = db.Column(db.String(100))
    total = db.Column(db.String(20))
    good_name = db.Column(db.String(100))
    spec = db.Column(db.String(100))
    unit = db.Column(db.String(20))
    quantity = db.Column(db.String(20))
    price = db.Column(db.String(20))
    payer = db.Column(db.String(50))
    stu_id = db.Column(db.String(50))
    bank_card = db.Column(db.String(50))
    folder_path = db.Column(db.String(200))

class InvoiceItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey('invoice.id', ondelete='CASCADE'), index=True)
    row = db.Column(db.Integer)
    name = db.Column(db.String(200))
    spec = db.Column(db.String(200))
    unit = db.Column(db.String(50))
    quantity = db.Column(db.String(50))
    price = db.Column(db.String(50))
    amount = db.Column(db.String(50))
    tax_rate = db.Column(db.String(50))
    tax = db.Column(db.String(50))
    invoice = db.relationship('Invoice', backref=db.backref('items', cascade='all, delete-orphan'))

with app.app_context():
    db.create_all()


def save_items_from_words(inv, words):
    """直接保存明细，并将金额转换为含税存入数据库"""
    if not words or not isinstance(words, dict):
        return

    def extract_list(key):
        v = words.get(key) or []
        if isinstance(v, list):
            return [el.get('word', '') if isinstance(el, dict) else str(el) for el in v]
        return [str(v)]

    names = extract_list('CommodityName')
    specs = extract_list('CommodityType')
    units = extract_list('CommodityUnit')
    nums = extract_list('CommodityNum')
    prices = extract_list('CommodityPrice') 
    amounts = extract_list('CommodityAmount') 
    taxes = extract_list('CommodityTax') 
    rates = extract_list('CommodityTaxRate')

    n = len(names)
    try:
        InvoiceItem.query.filter_by(invoice_id=inv.id).delete()
    except Exception: pass

    for i in range(n):
        try:
            raw_amt = float(amounts[i].replace(',', '')) if i < len(amounts) and amounts[i] else 0.0
            raw_tax = float(taxes[i].replace(',', '')) if i < len(taxes) and taxes[i] else 0.0
            # 数量清理
            q_str = nums[i].replace(',', '') if i < len(nums) and nums[i] else ''
            raw_qty = float(q_str) if q_str.replace('.','',1).isdigit() else 0.0
        except:
            raw_amt, raw_tax, raw_qty = 0.0, 0.0, 0.0

        total_with_tax = raw_amt + raw_tax
        
        # --- 关键逻辑：若无数量，单价 = 总额 ---
        if raw_qty != 0:
            final_price = total_with_tax / raw_qty
        else:
            final_price = total_with_tax

        item = InvoiceItem(
            invoice_id=inv.id,
            row=i + 1,
            name=names[i] if names[i] else '未知商品',
            spec=specs[i] if i < len(specs) else '',
            unit=units[i] if i < len(units) else '',
            quantity=str(raw_qty) if raw_qty != 0 else '',
            price=f"{final_price:.4f}",
            amount=f"{total_with_tax:.2f}",
            tax_rate=rates[i] if i < len(rates) else '',
            tax=str(raw_tax)
        )
        db.session.add(item)
    db.session.commit()
@app.route('/')
def index():
    # 使用 joinedload 预加载 items，减少数据库查询次数
    invoices = Invoice.query.options(db.joinedload(Invoice.items)).order_by(Invoice.id.desc()).all()
    
    for inv in invoices:
        inv.has_pay = False
        inv.has_order = False
        inv.files_list = []
        if inv.folder_path and os.path.exists(inv.folder_path):
            base_folder = os.path.basename(inv.folder_path)
            for f in os.listdir(inv.folder_path):
                if f == '.trash': continue
                if '支付' in f: inv.has_pay = True
                if '订单' in f: inv.has_order = True
                
                protected = False
                if f.startswith('发票') or f == f"{base_folder}.txt":
                    protected = True
                inv.files_list.append({'name': f, 'protected': protected})
                
    return render_template('index.html', invoices=invoices)

@app.route('/upload', methods=['POST'])
def upload():
    # 1. 获取文件列表支持批量上传
    files = request.files.getlist('invoice')
    if not files or all(f.filename == '' for f in files):
        flash('请选择发票文件进行上传', 'warning')
        return redirect(url_for('index'))
    
    aid = request.form.get('app_id') or BAIDU_CONFIG['APP_ID']
    ak = request.form.get('api_key') or BAIDU_CONFIG['API_KEY']
    sk = request.form.get('secret_key') or BAIDU_CONFIG['SECRET_KEY']
    client = AipOcr(aid, ak, sk)

    success_count = 0
    
    for file in files:
        if file.filename == '':
            continue

        # 临时文件使用随机数防止批量处理时同名冲突
        temp_filename = f"temp_{os.urandom(4).hex()}_{file.filename}"
        temp_path = os.path.join('storage', temp_filename)
        file.save(temp_path)
        
        try:
            # --- PDF 转图片 ---
            if file.filename.lower().endswith('.pdf'):
                images = convert_from_path(temp_path, dpi=200, poppler_path=POPPLER_PATH)
                buf = io.BytesIO()
                images[0].save(buf, format='JPEG', quality=85)
                image_data = buf.getvalue()
            else:
                with open(temp_path, 'rb') as f:
                    image_data = f.read()

            # --- 调用百度 OCR ---
            res = client.vatInvoice(image_data)
            if 'error_code' in res:
                flash(f"文件 {file.filename} 识别错误: {res.get('error_msg')}", 'danger')
                continue

            data = res.get('words_result', {})
            if not any([data.get('InvoiceCode'), data.get('InvoiceNum'), data.get('CommodityName')]):
                flash(f'文件 {file.filename} 识别失败：非标准发票', 'warning')
                continue

           # --- 提取信息 ---
            def extract_val(dct, key):
                v = dct.get(key)
                if isinstance(v, list): v = v[0] if v else None
                return v.get('word') if isinstance(v, dict) else v

            inv_num = extract_val(data, 'InvoiceNum')
            inv_code = extract_val(data, 'InvoiceCode')
            if inv_num: inv_num = inv_num.strip()
            if inv_code: inv_code = inv_code.strip()

            # ======= 新增：查重逻辑 =======
            if inv_num and inv_code:
                existing = Invoice.query.filter_by(inv_num=inv_num, inv_code=inv_code).first()
                if existing:
                    # 这里的提醒会显示在前端
                    flash(f'⚠️ 重复上传：发票号 {inv_num} 已存在，已自动跳过。', 'warning')
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    continue
            # =============================

            # 之前的 inv_num 赋默认值逻辑移动到查重之后，防止误判
            inv_num = inv_num or '未知号码'
            g_name = data.get('CommodityName', [{'word': '未知商品'}])[0]['word']
            safe_g_name = clean_path_name(g_name)
            payer = request.form.get('payer') or '匿名'

            # 1. 截取商品名前16位
            short_g_name = safe_g_name[:16]
            
            # 2. 仅截取发票号最后 4 位
            short_inv_num = inv_num[-4:] if len(inv_num) >= 4 else inv_num
            
            # 3. 组合名称：姓名_前16位商品名_发票后4位
            base_folder_name = f"{payer}_{short_g_name}_{short_inv_num}"
            inv_dir = os.path.join('storage', base_folder_name)
            
            # 核心拦截：如果文件夹已存在，说明该发票（或同名发票）已处理过
            if os.path.exists(inv_dir):
                flash(f'文件夹冲突：发票 {short_inv_num} 已存在，已自动跳过。', 'warning')
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                continue # 直接跳过，不执行后面的保存逻辑
            
            os.makedirs(inv_dir, exist_ok=True)
            final_folder_name = os.path.basename(inv_dir)

            # --- 数据库保存 ---
            new_inv = Invoice(
                inv_num=inv_num,
                inv_code=extract_val(data, 'InvoiceCode') or '',
                date=extract_val(data, 'InvoiceDate') or '',
                seller=extract_val(data, 'SellerName') or '',
                total=str(data.get('AmountInFiguers') or data.get('TotalAmount') or '0'),
                good_name=g_name,
                spec=extract_val(data, 'CommodityType') or '-',
                unit=extract_val(data, 'CommodityUnit') or '-',
                quantity=extract_val(data, 'CommodityNum') or '-',
                price=extract_val(data, 'CommodityPrice') or '-',
                payer=payer,
                stu_id=request.form.get('stu_id'),
                bank_card=request.form.get('bank_card'),
                folder_path=inv_dir
            )
            db.session.add(new_inv)
            db.session.flush()

            # --- 文件移动与TXT生成 ---
            ext = os.path.splitext(file.filename)[1]
            os.rename(temp_path, os.path.join(inv_dir, f"发票{ext}"))

            with open(os.path.join(inv_dir, f"{final_folder_name}.txt"), "w", encoding="utf-8") as f:
                f.write(f"姓名：{new_inv.payer}\n学号：{new_inv.stu_id}\n银行卡号：{new_inv.bank_card}")

            save_items_from_words(new_inv, data)
            db.session.commit()
            success_count += 1

        except Exception as e:
            db.session.rollback()
            flash(f'处理文件 {file.filename} 出错: {str(e)}', 'danger')
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    flash(f'成功批量处理 {success_count} 张发票', 'success')
    return redirect(url_for('index'))


@app.route('/delete_attachment/<int:inv_id>', methods=['POST'])
def delete_attachment(inv_id):
    """删除单个附件：表单需提供 `subfolder`（如 '订单截图' 或 '支付截图'，可为空）和 `filename`。"""
    inv = Invoice.query.get(inv_id)
    if not inv:
        return redirect(url_for('index'))

    # 现在所有附件都放在一级目录，忽略 subfolder 参数
    filename = request.form.get('filename', '').strip()
    if not filename:
        return redirect(url_for('index'))
    target = os.path.join(inv.folder_path, filename)

    # 不允许删除发票原件（文件名以 发票 开头）
    if filename.startswith('发票'):
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.accept_mimetypes.accept_json and not request.accept_mimetypes.accept_html
        if is_ajax:
            return jsonify({'ok': False, 'error': 'cannot_delete_invoice'})
        return redirect(url_for('index'))

    try:
        if os.path.exists(target):
            # 将文件移动到回收目录以支持撤销
            trash_dir = os.path.join(inv.folder_path, '.trash')
            os.makedirs(trash_dir, exist_ok=True)
            base = os.path.basename(target)
            import time
            trash_name = f"{int(time.time())}_{base}"
            trash_path = os.path.join(trash_dir, trash_name)
            shutil.move(target, trash_path)
            result_ok = True
        else:
            trash_name = None
            result_ok = False
    except Exception as e:
        result_ok = False
        # 记录错误但不阻塞用户操作
        try:
            with open(os.path.join(inv.folder_path, 'delete_attachment_error.log'), 'a', encoding='utf-8') as ef:
                ef.write(f"{filename}: {str(e)}\n")
        except Exception:
            pass

    # 如果是 AJAX 请求，返回 JSON，否则重定向回列表（兼容旧行为）
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.accept_mimetypes.accept_json and not request.accept_mimetypes.accept_html
    # 重新计算该发票当前的附件状态（是否存在支付/订单文件）
    has_pay = False
    has_order = False
    try:
        if os.path.exists(inv.folder_path):
            for f in os.listdir(inv.folder_path):
                if f == '.trash':
                    continue
                if '支付' in f:
                    has_pay = True
                if '订单' in f:
                    has_order = True
    except Exception:
        pass

    if is_ajax:
        return jsonify({'ok': result_ok, 'trash': (trash_name if result_ok else None), 'filename': filename, 'has_pay': has_pay, 'has_order': has_order})
    return redirect(url_for('index'))

@app.route('/delete/<int:inv_id>')
def delete_invoice(inv_id):
    inv = Invoice.query.get(inv_id)
    result_ok = False
    if inv:
        try:
            # 1. 物理删除文件夹
            if inv.folder_path and os.path.exists(inv.folder_path):
                import shutil
                shutil.rmtree(inv.folder_path)
            
            # 2. 数据库删除
            db.session.delete(inv)
            db.session.commit()
            result_ok = True
        except Exception as e:
            db.session.rollback()
            print(f"删除失败: {e}")

    # --- 关键修改：判断请求头 ---
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest' or \
              (request.accept_mimetypes.accept_json and not request.accept_mimetypes.accept_html)
    
    if is_ajax:
        return jsonify({'ok': result_ok}) # 如果是 JS 发起的请求，返回这个
    
    # 否则（手动输入链接）跳转回首页
    flash('删除成功' if result_ok else '记录不存在或删除失败')
    return redirect(url_for('index'))

@app.route('/restore_attachment/<int:inv_id>', methods=['POST'])
def restore_attachment(inv_id):
    inv = Invoice.query.get(inv_id)
    if not inv:
        return jsonify({'ok': False, 'error': 'not_found'})
    trash_name = request.form.get('trash')
    sub = ''
    filename = request.form.get('filename', '').strip()
    if not trash_name:
        return jsonify({'ok': False, 'error': 'missing_trash'})
    trash_path = os.path.join(inv.folder_path, '.trash', trash_name)
    if not os.path.exists(trash_path):
        return jsonify({'ok': False, 'error': 'trash_missing'})
    # 目标位置
    dst_dir = inv.folder_path
    os.makedirs(dst_dir, exist_ok=True)
    dst_path = os.path.join(dst_dir, filename)
    # 若目标已存在则尝试加后缀
    if os.path.exists(dst_path):
        name, ext = os.path.splitext(filename)
        import time
        filename = f"{name}_{int(time.time())}{ext}"
        dst_path = os.path.join(dst_dir, filename)
    try:
        shutil.move(trash_path, dst_path)
        return jsonify({'ok': True, 'filename': filename, 'sub': sub})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/preview_attachment/<int:inv_id>')
def preview_attachment(inv_id):
    inv = Invoice.query.get(inv_id)
    if not inv:
        return 'Not found', 404
    sub = request.args.get('subfolder', '').strip()
    filename = request.args.get('filename', '').strip()
    if sub:
        target = os.path.join(inv.folder_path, sub, filename)
    else:
        target = os.path.join(inv.folder_path, filename)
    if not os.path.exists(target):
        return 'Not found', 404
    # 安全：只返回 storage 下的文件
    # 直接返回文件流，让浏览器处理（图片/ pdf 等）
    mime, _ = mimetypes.guess_type(target)
    return send_file(target, mimetype=mime or 'application/octet-stream')


@app.route('/baidu_tutorial')
def baidu_tutorial():
    return render_template('baidu_tutorial.html')

@app.route('/download_all')
def download_all():
    invoices = Invoice.query.all()
    if not invoices:
        flash('当前没有任何发票记录，无法导出。', 'warning')
        return redirect(url_for('index'))
    
    data = []

    # helper: 格式化日期为 YYYY-MM-DD
    def fmt_date(s):
        if not s: return ''
        s = str(s).strip()
        from datetime import datetime
        patterns = ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y年%m月%d日', '%Y%m%d', '%Y-%m-%d %H:%M:%S']
        for p in patterns:
            try:
                dt = datetime.strptime(s, p)
                return dt.strftime('%Y-%m-%d')
            except: continue
        m = re.search(r'(20\d{2}[-/.年]?\d{1,2}[-/.月]?\d{1,2})', s)
        if m:
            candidate = m.group(1).replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
            parts = candidate.split('-')
            if len(parts) >= 3: return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
        return s

    for inv in invoices:
        if inv.items:
            # --- 情况 A：存在明细表 ---
            for it in inv.items:
                try:
                    # 转换数值用于判断
                    amt_val = float(it.amount) if it.amount else 0.0
                    qty_val = float(it.quantity) if it.quantity and float(it.quantity) != 0 else None
                    
                    # 核心逻辑：如果没有数量，单价等于总金额
                    if qty_val is None:
                        price_val = amt_val
                    else:
                        price_val = float(it.price) if it.price else (amt_val / qty_val)
                except:
                    # 如果转换出错，保留原始字符串
                    amt_val, price_val, qty_val = it.amount, it.price, it.quantity

                data.append({
                    "发票垫付人": inv.payer or '',
                    "学号": inv.stu_id or '',
                    "南京大学工行卡卡号": inv.bank_card or '',
                    "报销商品名称": it.name or inv.good_name or '',
                    "规格型号": it.spec or inv.spec or '',
                    "单位": it.unit or inv.unit or '',
                    "供应商": inv.seller or '',
                    "发票号": str(inv.inv_num or ''),
                    "发票代码": str(inv.inv_code or ''),
                    "数量": qty_val,
                    "总金额": amt_val,
                    "单价": price_val,
                    "开票日期": fmt_date(inv.date)
                })
        else:
            # --- 情况 B：无明细，使用主表汇总 ---
            try:
                total_val = float(inv.total.replace(',', '')) if inv.total else 0.0
                # 排除 '-' 或 '0' 等无效数量
                qty_str = str(inv.quantity).replace(',', '')
                qty_val = float(qty_str) if qty_str.replace('.','',1).isdigit() and float(qty_str) != 0 else None
                
                # 核心逻辑：如果没有数量，单价等于总金额
                if qty_val is None:
                    price_val = total_val
                else:
                    price_val = round(total_val / qty_val, 4)
            except:
                total_val, price_val, qty_val = inv.total, inv.price, inv.quantity

            data.append({
                "发票垫付人": inv.payer or '',
                "学号": inv.stu_id or '',
                "南京大学工行卡卡号": inv.bank_card or '',
                "报销商品名称": inv.good_name or '',
                "规格型号": inv.spec or '',
                "单位": inv.unit or '',
                "供应商": inv.seller or '',
                "发票号": str(inv.inv_num or ''),
                "发票代码": str(inv.inv_code or ''),
                "数量": qty_val,
                "总金额": total_val,
                "单价": price_val,
                "开票日期": fmt_date(inv.date)
            })

    # --- 后续 Excel 生成和 ZIP 打包逻辑 ---
    columns = ["发票垫付人", "学号", "南京大学工行卡卡号", "报销商品名称", "规格型号", "单位", "供应商", "发票号", "发票代码", "数量", "总金额", "单价", "开票日期"]
    df = pd.DataFrame(data, columns=columns)
    excel_p = "汇总.xlsx"
    df.to_excel(excel_p, index=False)
    
    zip_buf = io.BytesIO()
    root_folder = '报销材料汇总'
    with zipfile.ZipFile(zip_buf, 'w') as z:
        z.write(excel_p, arcname=os.path.join(root_folder, '汇总.xlsx'))
        
        missing_notes = []
        if os.path.exists('storage'):
            for inv in invoices:
                if not inv.folder_path or not os.path.exists(inv.folder_path): continue
                files = os.listdir(inv.folder_path)
                has_pay = any('支付' in f for f in files)
                has_order = any('订单' in f for f in files)
                if not (has_pay and has_order):
                    folder = os.path.basename(inv.folder_path)
                    missing_notes.append(f"{inv.payer} ({folder}) 缺少: {'支付 ' if not has_pay else ''}{'订单' if not has_order else ''}")
        
        if missing_notes:
            z.writestr(os.path.join(root_folder, '缺少附件提醒.txt'), '\n'.join(missing_notes))

        if os.path.exists('storage'):
            for folder in os.listdir('storage'):
                fpath = os.path.join('storage', folder)
                if os.path.isdir(fpath):
                    for fname in os.listdir(fpath):
                        if fname == '.trash': continue
                        z.write(os.path.join(fpath, fname), arcname=os.path.join(root_folder, folder, fname))
                        
    os.remove(excel_p)
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name="报销材料汇总.zip")

@app.route('/clear_all', methods=['POST'])
def clear_all():
    try:
        db.session.query(InvoiceItem).delete()
        db.session.query(Invoice).delete()
        db.session.commit()
        
        if os.path.exists('storage'):
            # 这种方式比直接 rmtree 更安全，能保留 storage 根目录
            for filename in os.listdir('storage'):
                file_path = os.path.join('storage', filename)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
        else:
            os.makedirs('storage', exist_ok=True)
            
        flash('已清空所有发票数据', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'清空失败: {str(e)}', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    os.makedirs('storage', exist_ok=True)
    app.run(debug=True)