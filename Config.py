import os
import torch.nn as nn




ConvFeatures = 50

pathInset = "autodl-fs"
train_Image_Path = os.path.join(os.getcwd(), pathInset, "Data", "copyTrain")
test_Image_Path = os.path.join(os.getcwd(), pathInset, "Data", "copyTest")
model_Image_Path = os.path.join(os.getcwd(), pathInset, "Data", "model")
excel_Path = os.path.join(os.getcwd(), pathInset, "Data", "copyData.xlsx")
sheet_Name = "Sheet"
excel_copyName = "Copy"
excel_modelColName = "Model"
excel_labelColName = "Z-Score"

batch_size = 32

model_name = "vgg16"
model_initWeight = True

# loss_function = nn.MSELoss()
loss_function = nn.L1Loss()

learningRate = 0.00001
lr_step = False
# EI：learningRate = 0.00001 发现收敛的很慢
# EI：learningRate = 0.001 梯度爆炸，输出一样
# EI：learningRate = 0.0003 梯度爆炸，输出一样
gamma = 0.9

epochs = 100

# debug时设置为0.0，正常模式设置为999
init_lowestLoss = 999



