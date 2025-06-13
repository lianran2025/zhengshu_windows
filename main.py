from flask_cors import CORS
from flask import Flask, request, send_file, jsonify
import os
from docx2pdf import convert
from PyPDF2 import PdfMerger
import uuid
import shutil
import logging
import win32com.client
import pythoncom
import zipfile

app = Flask(__name__)
CORS(app)  # 允许所有跨域请求，生产环境可指定 origins

UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'pdfs'
MERGED_FOLDER = 'merged'
COMPLETE_FOLDER = 'complete'  # 新增：完整压缩包文件夹
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(MERGED_FOLDER, exist_ok=True)
os.makedirs(COMPLETE_FOLDER, exist_ok=True)

# 全局任务进度字典
task_status = {}

# 新增：根路径路由，避免404错误
@app.route('/', methods=['GET'])
def index():
    return jsonify({
        'status': 'running',
        'message': '证书生成服务正在运行',
        'version': '2.0',
        'endpoints': [
            'POST /upload - 上传文件',
            'POST /convert/<task_id> - 转换PDF',
            'POST /merge/<task_id> - 合并PDF',
            'POST /package/<task_id> - 生成完整压缩包',
            'GET /progress/<task_id> - 查询进度',
            'GET /download/<task_id>/<filetype> - 下载文件'
        ]
    })

# 新增：健康检查接口
@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'healthy',
        'timestamp': str(uuid.uuid4()),
        'active_tasks': len(task_status)
    })

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error': '没有上传文件'}), 400
        
        task_id = str(uuid.uuid4())
        task_folder = os.path.join(UPLOAD_FOLDER, task_id)
        os.makedirs(task_folder, exist_ok=True)
        
        uploaded_count = 0
        for file in files:
            if file.filename:
                file.save(os.path.join(task_folder, file.filename))
                uploaded_count += 1
        
        print(f"上传成功，任务ID: {task_id}，文件数量: {uploaded_count}")
        return jsonify({
            'task_id': task_id, 
            'msg': '上传成功',
            'file_count': uploaded_count
        })
    except Exception as e:
        print(f"上传文件失败: {str(e)}")
        return jsonify({'error': f'上传失败: {str(e)}'}), 500

@app.route('/convert/<task_id>', methods=['POST'])
def convert_to_pdf(task_id):
    try:
        task_folder = os.path.join(UPLOAD_FOLDER, task_id)
        if not os.path.exists(task_folder):
            return jsonify({'error': '任务文件夹不存在'}), 404
            
        pdf_task_folder = os.path.join(PDF_FOLDER, task_id)
        os.makedirs(pdf_task_folder, exist_ok=True)
        files = [f for f in os.listdir(task_folder) if f.endswith('.docx') and not f.startswith('~$')]
        total = len(files)
        
        if total == 0:
            return jsonify({'error': '没有找到可转换的docx文件'}), 400
        
        # 初始化进度，新增 logs 字段
        task_status[task_id] = {
            'total': total,
            'current': 0,
            'current_file': '',
            'results': [],
            'done': False,
            'convert_done': False,
            'merge_done': False,
            'package_done': False,
            'logs': []  # 新增
        }
        
        log = f"开始批量转换PDF，任务ID: {task_id}，文件数量: {total}"
        print(log)
        task_status[task_id]['logs'].append(log)
        
        # 调用批量转换函数
        results = batch_convert_docx_to_pdf(task_folder, pdf_task_folder, task_id)
        
        # 标记转换完成
        task_status[task_id]['convert_done'] = True
        task_status[task_id]['done'] = False  # 整体还未完成
        log = f"PDF批量转换完成，任务ID: {task_id}"
        print(log)
        task_status[task_id]['logs'].append(log)
        
        return jsonify({
            'msg': '转换完成', 
            'results': results, 
            'pdf_folder': pdf_task_folder,
            'success_count': len([r for r in results if r['status'] == 'success'])
        })
    except Exception as e:
        print(f"转换PDF失败: {str(e)}")
        if task_id in task_status:
            task_status[task_id]['logs'].append(f"转换PDF失败: {str(e)}")
        return jsonify({'error': f'转换失败: {str(e)}'}), 500

# 新增：批量转换函数，复用 Word 进程
def batch_convert_docx_to_pdf(docx_folder, pdf_folder, task_id=None):
    results = []
    word = None
    pythoncom.CoInitialize()  # 新增：初始化 COM
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        files = [f for f in os.listdir(docx_folder) if f.endswith('.docx') and not f.startswith('~$')]
        total = len(files)
        for idx, filename in enumerate(files):
            src = os.path.abspath(os.path.join(docx_folder, filename))
            dst = os.path.abspath(os.path.join(pdf_folder, filename.replace('.docx', '.pdf')))
            try:
                doc = word.Documents.Open(src)
                doc.SaveAs(dst, FileFormat=17)
                doc.Close()
                result = {'file': filename, 'status': 'success'}
                log = f"转换成功: {filename}"
                print(log)
            except Exception as e:
                result = {'file': filename, 'status': 'fail', 'reason': str(e)}
                log = f"转换失败: {filename} - {str(e)}"
                print(log)
            results.append(result)
            # 实时更新进度和日志
            if task_id and task_id in task_status:
                task_status[task_id]['current'] = idx + 1
                task_status[task_id]['current_file'] = filename
                task_status[task_id]['results'] = results.copy()
                if 'logs' in task_status[task_id]:
                    task_status[task_id]['logs'].append(log)
        return results
    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()  # 新增：释放 COM

@app.route('/progress/<task_id>', methods=['GET'])
def get_progress(task_id):
    status = task_status.get(task_id)
    if not status:
        return jsonify({'error': '任务不存在'}), 404
    
    # 计算整体完成状态
    all_done = status.get('convert_done', False) and status.get('merge_done', False) and status.get('package_done', False)
    status['done'] = all_done
    
    # 确保返回 logs 字段
    if 'logs' not in status:
        status['logs'] = []
    return jsonify(status)

@app.route('/merge/<task_id>', methods=['POST'])
def merge_pdfs(task_id):
    try:
        pdf_task_folder = os.path.join(PDF_FOLDER, task_id)
        if not os.path.exists(pdf_task_folder):
            return jsonify({'error': 'PDF文件夹不存在'}), 404
            
        merged_file = os.path.join(MERGED_FOLDER, f'{task_id}_merged.pdf')
        merger = PdfMerger()
        pdfs = [f for f in sorted(os.listdir(pdf_task_folder)) if f.endswith('.pdf')]
        
        log = f"合并PDF，任务ID: {task_id}，PDF数量: {len(pdfs)}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        
        if not pdfs:
            log = "没有可合并的 PDF 文件"
            print(log)
            if task_id in task_status and 'logs' in task_status[task_id]:
                task_status[task_id]['logs'].append(log)
            return jsonify({'msg': log, 'merged_file': None}), 400
        
        for filename in pdfs:
            pdf_path = os.path.join(pdf_task_folder, filename)
            log = f"合并文件: {filename}"
            print(log)
            if task_id in task_status and 'logs' in task_status[task_id]:
                task_status[task_id]['logs'].append(log)
            try:
                merger.append(pdf_path)
            except Exception as e:
                log = f"合并文件失败 {filename}: {str(e)}"
                print(log)
                if task_id in task_status and 'logs' in task_status[task_id]:
                    task_status[task_id]['logs'].append(log)
                continue
        
        merger.write(merged_file)
        merger.close()
        log = f"合并完成，输出文件: {merged_file}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        
        # 更新任务状态
        if task_id in task_status:
            task_status[task_id]['merge_done'] = True
        
        return jsonify({
            'msg': '合并完成', 
            'merged_file': merged_file,
            'pdf_count': len(pdfs)
        })
    except Exception as e:
        log = f"合并PDF失败: {str(e)}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        return jsonify({'error': f'合并失败: {str(e)}'}), 500

# 新增：生成完整压缩包接口
@app.route('/package/<task_id>', methods=['POST'])
def package_complete_files(task_id):
    try:
        data = request.get_json() or {}
        filename = data.get('filename', f'certificates_{task_id}.zip')
        
        # 从文件名中提取文件夹名称（去掉.zip后缀）
        folder_name = filename.replace('.zip', '') if filename.endswith('.zip') else filename
        
        # 获取任务相关的文件路径
        docx_folder = os.path.join(UPLOAD_FOLDER, task_id)
        merged_pdf = os.path.join(MERGED_FOLDER, f'{task_id}_merged.pdf')
        complete_zip_path = os.path.join(COMPLETE_FOLDER, f'{task_id}_{filename}')
        
        log = f"开始生成完整压缩包，任务ID: {task_id}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"文件夹名称: {folder_name}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"docx文件夹: {docx_folder}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"合并PDF: {merged_pdf}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"输出路径: {complete_zip_path}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        
        # 检查必要文件是否存在
        if not os.path.exists(docx_folder):
            return jsonify({'error': 'docx文件夹不存在'}), 404
        if not os.path.exists(merged_pdf):
            return jsonify({'error': '合并PDF文件不存在'}), 404
        
        # 创建完整压缩包，所有文件都放在指定文件夹内
        file_count = 0
        with zipfile.ZipFile(complete_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 添加所有docx文件到文件夹内
            if os.path.exists(docx_folder):
                for file in os.listdir(docx_folder):
                    if file.endswith('.docx') and not file.startswith('~$'):
                        file_path = os.path.join(docx_folder, file)
                        # 将文件放在文件夹内：文件夹名/文件名
                        archive_path = f"{folder_name}/{file}"
                        zipf.write(file_path, archive_path)
                        log = f"添加docx文件到文件夹: {archive_path}"
                        print(log)
                        if task_id in task_status and 'logs' in task_status[task_id]:
                            task_status[task_id]['logs'].append(log)
                        file_count += 1
            
            # 添加合并后的PDF文件到文件夹内
            if os.path.exists(merged_pdf):
                # 将PDF文件放在文件夹内：文件夹名/合并证书.pdf
                archive_path = f"{folder_name}/合并证书.pdf"
                zipf.write(merged_pdf, archive_path)
                log = f"添加合并PDF文件到文件夹: {archive_path}"
                print(log)
                if task_id in task_status and 'logs' in task_status[task_id]:
                    task_status[task_id]['logs'].append(log)
                file_count += 1
        
        # 更新任务状态
        if task_id in task_status:
            task_status[task_id]['package_done'] = True
            task_status[task_id]['complete_zip_path'] = complete_zip_path
            task_status[task_id]['folder_name'] = folder_name
        
        log = f"完整压缩包生成成功: {complete_zip_path}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"解压后将创建文件夹: {folder_name}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        log = f"包含文件数量: {file_count}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        
        return jsonify({
            'status': 'success',
            'message': '完整压缩包生成成功',
            'filename': filename,
            'folder_name': folder_name,
            'file_count': file_count
        })
        
    except Exception as e:
        log = f"生成完整压缩包失败: {str(e)}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        return jsonify({
            'status': 'error',
            'message': f'生成完整压缩包失败: {str(e)}'
        }), 500

@app.route('/download/<task_id>/<filetype>', methods=['GET'])
def download_file(task_id, filetype):
    try:
        if filetype == 'merged':
            file_path = os.path.join(MERGED_FOLDER, f'{task_id}_merged.pdf')
        elif filetype == 'pdfs':
            file_path = os.path.join(PDF_FOLDER, task_id)
            # 可打包为 zip 返回
            shutil.make_archive(file_path, 'zip', file_path)
            file_path += '.zip'
        elif filetype == 'docx':
            file_path = os.path.join(UPLOAD_FOLDER, task_id)
            shutil.make_archive(file_path, 'zip', file_path)
            file_path += '.zip'
        elif filetype == 'complete':
            # 新增：下载完整压缩包
            filename = request.args.get('filename', f'certificates_{task_id}.zip')
            file_path = os.path.join(COMPLETE_FOLDER, f'{task_id}_{filename}')
            
            if not os.path.exists(file_path):
                return jsonify({
                    'status': 'error',
                    'message': '完整压缩包不存在，请稍后重试'
                }), 404
            
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/zip'
            )
        else:
            return jsonify({'error': '文件类型错误'}), 400
        
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
            
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        log = f"下载文件失败: {str(e)}"
        print(log)
        if task_id in task_status and 'logs' in task_status[task_id]:
            task_status[task_id]['logs'].append(log)
        return jsonify({'error': f'下载失败: {str(e)}'}), 500

if __name__ == '__main__':
    print("=" * 50)
    print("证书生成服务启动中...")
    print("服务地址: http://0.0.0.0:5000")
    print("健康检查: http://0.0.0.0:5000/health")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)