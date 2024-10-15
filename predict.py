import os
import json

import torch
from PIL import Image
from torchvision import transforms
import pandas as pd
import matplotlib.pyplot as plt

from ConvModel import vgg


import time

print(os.getcwd())
def main():
    device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
    print(device)

    excelPath = "../Data/copyData1.xlsx" #这里是excel路径
    print(os.path.isfile(excelPath))
    df = pd.read_excel(excelPath, sheet_name="Sheet")



    testPath = "../Data/copyTest1" #这里是测试集路径
    testImages = os.listdir(testPath)

    modelPath = "../Data/model" # 这里是书法家原作路径

    data_transform = transforms.Compose([transforms.ToTensor()])


    # create model
    model = vgg(model_name="vgg16").to(device)
    # model = densenet121().to(device)
    # model = resnet101(50, True).to(device)
    # load model weights
    weights_path = "../vggpth/vgg16_Sociability.pth" # 这里加载pth
    assert os.path.exists(weights_path), "file: '{}' dose not exist.".format(weights_path)
    model.load_state_dict(torch.load(weights_path, map_location=device))

    count = 0
    start_time = time.time()
    totaldiff = 0.0

    model.eval()
    with torch.no_grad():
        # predict class

        for image in testImages:
            img_copy_path = os.path.join(testPath, image)
            row_index = df[df['Copy'].str.contains(image, na=False)].index.min()
            score = df.loc[row_index, 'Z-EI.Sociability'] # 这里修改excel标签
            modelPic = df.loc[row_index, 'Model']
            img_model_path = os.path.join(modelPath, modelPic)

            # load image
            assert os.path.exists(img_copy_path), "file: '{}' dose not exist.".format(img_copy_path)
            img_copy = Image.open(img_copy_path)
            img_copy = data_transform(img_copy)

            assert os.path.exists(img_model_path), "file: '{}' dose not exist.".format(img_model_path)
            img_model = Image.open(img_model_path)
            img_model = data_transform(img_model)

            img_copy = img_copy.unsqueeze(0)  # 假设 img_copy 是三维的
            img_model = img_model.unsqueeze(0)  # 如果 img_model 也是图像数据

            features, outputs = model(img_copy.to(device), img_model.to(device))

	# 这里的outputs 就是最终预测的值，我后面代码的作用是，把预测的值写入excel里面，这一步你可以忽略，后面的代码你可以删除相关的excel操作，我来不及改了

            df.loc[row_index, 'pre-Sociability'] = outputs.item()
            # diff, output = torch.squeeze(model(img1.to(device), img2.to(device))).cpu()
            count += 1
            totaldiff += abs(outputs.item() - score)

            print(count, outputs, outputs.item())
        df.to_excel(excelPath, sheet_name="Sheet", index=False)
        print('时间为:',time.time()-start_time)
        print('mae:',totaldiff / count)






if __name__ == '__main__':
    main()
