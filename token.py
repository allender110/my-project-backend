import os
import json
from datetime import datetime
import pandas as pd
from flask import Flask, request, jsonify
import openpyxl
from flask_cors import CORS

app = Flask(__name__, static_url_path='/static')
CORS(app)

# 硬编码的 Excel 和模型图像路径
excel_path = "./copyData1.xlsx"  # 固定的 Excel 文件路径
modelPath = "./model"  # 固定的模型图像路径

# 读取 Excel 数据（只需要加载一次）
df = pd.read_excel(excel_path, sheet_name="Sheet")

# 保存上传的图片的目录
upload_folder = './images'
os.makedirs(upload_folder, exist_ok=True)

# 初始化或读取Excel表
def save_scores_to_excel(filename, image_name, scores):
    if not os.path.exists(filename):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # 添加列标题
        sheet.append(["ImageName", "Timestamp", "美观", "情绪", "自控", "情感", "社交", "EI"])
    else:
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active

    # 获取当前时间
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 将新行添加到表格
    sheet.append([
        image_name,
        current_time,
        scores['美观'],
        scores['情绪'],
        scores['自控'],
        scores['情感'],
        scores['社交'],
        scores['EI']
    ])

    workbook.save(filename)

@app.route('/predict', methods=['POST'])
def predict():
    try:
        # 从请求中获取上传的测试图像
        test_image = request.files['image']
        filename = test_image.filename

        # 保存上传的图像到 images 文件夹
        img_copy_path = os.path.join(upload_folder, filename)
        test_image.save(img_copy_path)

        # 根据测试图像文件名查找 Excel 中的分数
        row_index = df[df['Copy'].str.contains(test_image.filename, na=False)].index.min()
        # 从Excel中提取对应模型的score
        scores = {
            '美观': df.loc[row_index, 'Z-Score'],
            '情绪': df.loc[row_index, 'Z-EI'],
            '自控': df.loc[row_index, 'Z-EI.Self-Control'],
            '情感': df.loc[row_index, 'Z-EI.Emotionality'],
            '社交': df.loc[row_index, 'Z-EI.Sociability'],
            'EI': df.loc[row_index, 'Z-EI']
        }

        # 在这之后添加数据到Excel中
        save_scores_to_excel('data.xlsx', filename, scores)
        return jsonify(scores)

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/history', methods=['GET'])
def get_history():
    try:
        if os.path.exists('data.xlsx'):
            df = pd.read_excel('data.xlsx')
            records = df.to_dict(orient='records')
            return jsonify(records)
        else:
            return jsonify([])

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/images/<path:filename>')
def serve_image(filename):
    return send_from_directory(upload_folder, filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
