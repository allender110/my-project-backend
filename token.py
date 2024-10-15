import os
import time
import json
from datetime import datetime

import torch
from PIL import Image
from torchvision import transforms
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory

import openpyxl

from ConvModel import vgg
from flask_cors import CORS
# 假设图片保存在 './images' 目录下
app = Flask(__name__, static_url_path='/static')
# app = Flask(__name__)
CORS(app)

# 模型初始化
device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")

# 加载六个不同的VGG模型
models = []
weights_paths = ["../vggpth/vgg16_Score.pth", "../vggpth/vgg16_EI.pth", "../vggpth/vgg16_Wellbeing.pth", "../vggpth/vgg16_SelfControl.pth", "../vggpth/vgg16_Emotionality.pth", "../vggpth/vgg16_Sociability.pth"]
for weights_path in weights_paths:
    model = vgg(model_name="vgg16").to(device)
    assert os.path.exists(weights_path), "file: '{}' does not exist.".format(weights_path)
    model.load_state_dict(torch.load(weights_path, map_location=device))
    model.eval()
    models.append(model)

# 数据预处理
data_transform = transforms.Compose([transforms.ToTensor()])

# 硬编码的 Excel 和模型图像路径
excel_path = "../Data/copyData1.xlsx"  # 固定的 Excel 文件路径
modelPath = "../Data/model"  # 固定的模型图像路径

# 读取 Excel 数据和模型图像（只需要加载一次）
df = pd.read_excel(excel_path, sheet_name="Sheet")

# 保存上传的图片的目录
upload_folder = './images'
os.makedirs(upload_folder, exist_ok=True)


# 初始化或读取Excel表

# 初始化或读取Excel表
def save_scores_to_excel(filename, image_name, scores):
    # 如果文件不存在，则创建一个新的工作簿
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
        scores = {}  # 创建一个字典来存储每个模型的得分
        scores[0] = df.loc[row_index, 'Z-Score']
        scores[1] = df.loc[row_index, 'Z-EI']
        scores[2] = df.loc[row_index, 'Z-EI.Well-being']
        scores[3] = df.loc[row_index, 'Z-EI.Self-Control']
        scores[4] = df.loc[row_index, 'Z-EI.Emotionality']
        scores[5] = df.loc[row_index, 'Z-EI.Sociability']
        modelPic = df.loc[row_index, 'Model']
        img_model_path = os.path.join(modelPath, modelPic)

        # load image
        assert os.path.exists(img_copy_path), "file: '{}' dose not exist.".format(img_copy_path)
        img_copy = Image.open(img_copy_path)
        img_copy = data_transform(img_copy)

        assert os.path.exists(img_model_path), "file: '{}' dose not exist.".format(img_model_path)
        img_model = Image.open(img_model_path)
        img_model = data_transform(img_model)

        # 进行预测
        # 进行预测
        results = {}
        outputs={}
        features={}
        with torch.no_grad():
            for i, model in enumerate(models):
            # 从Excel中提取对应模型的score
                features[i], outputs[i]= models[i](img_copy.to(device), img_model.to(device))  # 假设模型返回预测和特征
            results['美观'] = outputs[0].item()
            results['情绪'] = outputs[2].item()
            results['自控'] = outputs[3].item()
            results['情感'] = outputs[4].item()
            results['社交'] = outputs[5].item()
            results['EI'] = outputs[1].item()

        # 在这之后添加数据到Excel中
        save_scores_to_excel('data.xlsx',filename, results)
        return jsonify(results)



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




# 用于展示图片的API
@app.route('/images/<path:filename>')
def serve_image(filename):
    return send_from_directory(upload_folder, filename)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)




