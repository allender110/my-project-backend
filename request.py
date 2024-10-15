import requests

# 设置要上传的图片路径
image_path = "../Data/copyTest/1_1157230101_yanzhenqing_34_5.png"

# 设置 API 地址
url = "http://localhost:5000/predict"

# 打开测试图像并准备发送请求
with open(image_path, 'rb') as img:
    files = {'image': img}  # 构建要发送的文件
    try:
        # 发送 POST 请求，并上传图像
        response = requests.post(url, files=files)

        # 检查请求是否成功
        if response.status_code == 200:
            # 解析并打印返回的 JSON 数据
            result = response.json()
            print(f"美观得分: {result['美观得分']}")
            print(f"EI总分: {result['EI总分']}")
            print(f"情绪得分: {result['情绪得分']}")
            print(f"自控得分: {result['自控得分']}")
            print(f"情感得分: {result['情感得分']}")
            print(f"社交得分: {result['社交得分']}")

        else:
            print(f"Error: {response.status_code}")
            print(response.text)
    except Exception as e:
        print(f"An error occurred: {str(e)}")


