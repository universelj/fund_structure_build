#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基金结构图Web应用 - Draw.io生成工具
整合了 start_app.py 的依赖检查与启动逻辑
"""

import os
import sys
import io
import re
import time
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename, safe_join

# 导入转换器（从顶层 generator.py 导入）
from generator import ExcelToDrawioGenerator, InputValidationError

# ──────────────────── Flask 应用配置 ────────────────────

app = Flask(__name__)

def _env_int(name, default):
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return int(default)


app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['MAX_XML_BYTES'] = _env_int('MAX_XML_BYTES', 0)  # 0 表示不限制

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
OUTPUT_TTL_HOURS = _env_int('OUTPUT_TTL_HOURS', 0)
OUTPUT_CLEANUP_INTERVAL_SECONDS = _env_int('OUTPUT_CLEANUP_INTERVAL_SECONDS', 900)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True)


# ──────────────────── 辅助函数 ────────────────────

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def sanitize_filename(filename):
    """清理文件名，移除特殊字符"""
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    filename = re.sub(r'[\r\n\t]', ' ', filename)
    filename = filename.strip()
    if len(filename) > 100:
        filename = filename[:100]
    return filename


def safe_path(folder, filename):
    """安全路径拼接，防止路径遍历攻击（兼容中文文件名）"""
    if not filename:
        return None
    if '..' in filename or '/' in filename or '\\' in filename:
        return None
    try:
        base = os.path.abspath(folder)
        joined = safe_join(base, filename)
    except Exception:
        return None
    return joined


_last_cleanup = 0


def cleanup_outputs(ttl_hours):
    """清理过期输出文件"""
    if ttl_hours <= 0:
        return
    cutoff = time.time() - ttl_hours * 3600
    for fname in os.listdir(OUTPUT_FOLDER):
        if not fname.endswith('.drawio'):
            continue
        fpath = os.path.join(OUTPUT_FOLDER, fname)
        try:
            if os.path.getmtime(fpath) < cutoff:
                os.remove(fpath)
        except OSError:
            pass


def maybe_cleanup_outputs():
    global _last_cleanup
    if OUTPUT_TTL_HOURS <= 0:
        return
    now = time.time()
    if now - _last_cleanup < OUTPUT_CLEANUP_INTERVAL_SECONDS:
        return
    _last_cleanup = now
    cleanup_outputs(OUTPUT_TTL_HOURS)


# ──────────────────── 路由 ────────────────────

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """上传Excel文件并生成Draw.io文件"""
    try:
        maybe_cleanup_outputs()
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            if not filename:
                return jsonify({'error': '文件名不合法'}), 400
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            temp_filename = f"{timestamp}_{filename}"
            filepath = os.path.join(UPLOAD_FOLDER, temp_filename)
            file.save(filepath)

            response = None
            status_code = 200
            try:
                generator = ExcelToDrawioGenerator()
                fund_info, nodes, positions, fund_id = generator.generate_from_excel(filepath)

                chart_title = fund_info.get('chart_title', '基金结构图')
                safe_title = sanitize_filename(chart_title).strip()
                if not safe_title:
                    safe_title = '基金结构图'

                drawio_filename = f"{safe_title}.drawio"
                drawio_filepath = os.path.join(OUTPUT_FOLDER, drawio_filename)

                if os.path.exists(drawio_filepath):
                    name, ext = os.path.splitext(drawio_filename)
                    drawio_filename = f"{name}_{timestamp}{ext}"
                    drawio_filepath = os.path.join(OUTPUT_FOLDER, drawio_filename)

                generator.generate_drawio_xml(fund_info, nodes, positions, fund_id, drawio_filepath)

                response = jsonify({
                    'success': True,
                    'drawio_file': drawio_filename,
                    'fund_name': fund_info.get('fund_name', ''),
                    'chart_title': chart_title,
                    'investor_count': len([n for n in nodes.values() if n['category'] == 'investor'])
                })
            except InputValidationError as e:
                response = jsonify({'error': str(e)})
                status_code = 400
            except Exception as e:
                response = jsonify({'error': f'生成Draw.io文件失败: {str(e)}'})
                status_code = 500
            finally:
                try:
                    os.remove(filepath)
                except OSError:
                    pass

            return response, status_code

        return jsonify({'error': '不支持的文件格式，请上传Excel文件'}), 400

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """下载Draw.io文件"""
    try:
        file_path = safe_path(OUTPUT_FOLDER, filename)
        if file_path is None or not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404

        return send_file(file_path, as_attachment=True, download_name=filename)

    except Exception as e:
        return jsonify({'error': f'下载文件时出错: {str(e)}'}), 500


@app.route('/api/list_files')
def list_files():
    """列出所有生成的Draw.io文件"""
    try:
        maybe_cleanup_outputs()
        files = []
        for fname in os.listdir(OUTPUT_FOLDER):
            if fname.endswith('.drawio'):
                fpath = os.path.join(OUTPUT_FOLDER, fname)
                stat = os.stat(fpath)
                files.append({
                    'filename': fname,
                    'size': stat.st_size,
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                })

        files.sort(key=lambda x: x['modified'], reverse=True)
        return jsonify({'files': files})

    except Exception as e:
        return jsonify({'error': f'获取文件列表时出错: {str(e)}'}), 500


@app.route('/api/delete/<filename>', methods=['DELETE'])
def delete_file(filename):
    """删除文件"""
    try:
        file_path = safe_path(OUTPUT_FOLDER, filename)
        if file_path is None or not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404

        os.remove(file_path)
        return jsonify({'success': True, 'message': '文件删除成功'})

    except Exception as e:
        return jsonify({'error': f'删除文件时出错: {str(e)}'}), 500


@app.route('/api/get_xml/<filename>')
def get_xml(filename):
    """获取Draw.io文件的XML内容（供嵌入编辑器加载）"""
    try:
        file_path = safe_path(OUTPUT_FOLDER, filename)
        if file_path is None or not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404

        with open(file_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()
        return jsonify({'xml': xml_content, 'filename': filename})

    except Exception as e:
        return jsonify({'error': f'读取文件时出错: {str(e)}'}), 500


@app.route('/api/save_xml/<filename>', methods=['POST'])
def save_xml(filename):
    """保存编辑后的Draw.io XML"""
    try:
        file_path = safe_path(OUTPUT_FOLDER, filename)
        if file_path is None:
            return jsonify({'error': '非法路径'}), 400

        data = request.get_json()
        if not data or 'xml' not in data:
            return jsonify({'error': '缺少XML数据'}), 400

        xml_data = data['xml']
        if not isinstance(xml_data, str):
            return jsonify({'error': 'XML数据格式错误'}), 400
        max_bytes = app.config.get('MAX_XML_BYTES', 0)
        if max_bytes > 0 and len(xml_data.encode('utf-8')) > max_bytes:
            return jsonify({'error': 'XML内容过大，请减少图形复杂度后重试'}), 413

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(xml_data)
        return jsonify({'success': True, 'message': '保存成功'})

    except Exception as e:
        return jsonify({'error': f'保存文件时出错: {str(e)}'}), 500

# ──────────────────── 启动入口（整合自 start_app.py） ────────────────────

def check_dependencies():
    """检查依赖包"""
    required = ['flask', 'pandas', 'openpyxl', 'werkzeug', 'xlrd']
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("[X] 缺少以下依赖包:")
        for pkg in missing:
            print(f"   - {pkg}")
        print(f"\n请运行: pip install {' '.join(missing)}")
        return False
    return True


def check_key_files():
    """检查关键文件"""
    key_files = ['generator.py', 'templates/base.html', 'templates/index.html']
    missing = [f for f in key_files if not os.path.exists(f)]
    if missing:
        print("[X] 缺少关键文件:")
        for f in missing:
            print(f"   - {f}")
        return False
    return True


if __name__ == '__main__':
    # 设置标准输出编码为 UTF-8
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

    print("=== 基金结构图生成 ===")
    print("=" * 50)

    if not check_dependencies():
        sys.exit(1)
    print("[OK] 依赖检查通过")

    if not check_key_files():
        sys.exit(1)
    print("[OK] 文件检查通过")

    # 创建必要的文件夹
    for folder in ['uploads', 'outputs', 'static']:
        os.makedirs(folder, exist_ok=True)
    print("[OK] 文件夹创建完成")

    maybe_cleanup_outputs()

    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'

    print("\n" + "=" * 60)
    print("  基金结构图Web应用启动中...")
    print("  功能：Excel -> Draw.io 格式转换")
    print("=" * 60)
    print(f"  上传文件夹: {UPLOAD_FOLDER}")
    print(f"  输出文件夹: {OUTPUT_FOLDER}")
    print(f"  访问地址: http://localhost:5001")
    print("  支持格式: .drawio (可在 https://app.diagrams.net 中打开)")
    print("  按 Ctrl+C 停止应用")
    print("=" * 60 + "\n")

    app.run(debug=debug, host='0.0.0.0', port=5001)
