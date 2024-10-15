import torch.nn as nn
import torch
import Config as cfg

# 官方模型的地址
model_urls = {
    'vgg11': 'https://download.pytorch.org/models/vgg11-bbd30ac9.pth',
    'vgg13': 'https://download.pytorch.org/models/vgg13-c768596a.pth',
    'vgg16': 'https://download.pytorch.org/models/vgg16-397923af.pth',
    'vgg19': 'https://download.pytorch.org/models/vgg19-dcbb9e9d.pth'
}

class ConvVGG(nn.Module):
    def __init__(self, features, init_weights=True):
        super(ConvVGG, self).__init__()
        self.features = features  # 返回了提取特征网络
        self.getDeepFeatures = nn.Sequential(

            nn.Linear(512 * 7 * 7, 4096),
            nn.ReLU(True),
            nn.Dropout(p=0.5),
            nn.Linear(4096, 4096),
            nn.ReLU(True),
            nn.Dropout(p=0.5),
            nn.Linear(4096, cfg.ConvFeatures)
            # nn.ReLU(True),
            # nn.Dropout(p=0.5),
            # nn.Linear(cfg.ConvFeatures, cfg.ConvFeatures)
        )

        self.predict = nn.Sequential(
            nn.Linear(cfg.ConvFeatures, cfg.ConvFeatures),
            nn.Linear(cfg.ConvFeatures, 1),
        )


        if init_weights:
            self._initialize_weights()

    def forward(self, img_copy, img_model):
        img_copy = self.features(img_copy)
        img_copy = torch.flatten(img_copy, start_dim=0)
        img_copy = self.getDeepFeatures(img_copy)

        img_model = self.features(img_model)
        img_model = torch.flatten(img_model, start_dim=0)
        img_model = self.getDeepFeatures(img_model)

        # img_copy = torch.flatten(img_copy, start_dim=1)
        # img_model = torch.flatten(img_model, start_dim=1)
        featuresDiff = img_copy - img_model
        # features = torch.cat((img_copy, img_model), dim = 1)
        result = self.predict(featuresDiff)
        return featuresDiff, result

    # 初始化参数
    # 遍历每一个模块，使用xavier进行初始化
    def _initialize_weights(self):
        for m in self.modules():
            if isinstance(m, nn.Conv2d):
                # nn.init.kaiming_normal_(m.weight, mode='fan_out', nonlinearity='relu')
                nn.init.xavier_uniform_(m.weight)
                # 如果有偏置，置偏置为0
                if m.bias is not None:
                    nn.init.constant_(m.bias, 0)
            elif isinstance(m, nn.Linear):
                nn.init.xavier_uniform_(m.weight)
                # nn.init.normal_(m.weight, 0, 0.01)
                nn.init.constant_(m.bias, 0)


# 调用这个函数，输一个vgg的模型，就会将所有卷积层和池化层的流程返回
def make_features(convCfg: list):
    layers = []
    in_channels = 1
    for v in convCfg:
        if v == "M":
            layers += [nn.MaxPool2d(kernel_size=2, stride=2)]
        else:
            conv2d = nn.Conv2d(in_channels, v, kernel_size=3, padding=1)
            layers += [conv2d, nn.ReLU(True)]
            in_channels = v
    return nn.Sequential(*layers)


# 预先设置了每个模型的流程
cfgs = {
    'vgg11': [64, 'M', 128, 'M', 256, 256, 'M', 512, 512, 'M', 512, 512, 'M'],
    'vgg13': [64, 64, 'M', 128, 128, 'M', 256, 256, 'M', 512, 512, 'M', 512, 512, 'M'],
    'vgg16': [64, 64, 'M', 128, 128, 'M', 256, 256, 256, 'M', 512, 512, 512, 'M', 512, 512, 512, 'M'],
    'vgg19': [64, 64, 'M', 128, 128, 'M', 256, 256, 256, 256, 'M', 512, 512, 512, 512, 'M', 512, 512, 512, 512, 'M'],
}


def vgg(model_name="vgg16", **kwargs):
    assert model_name in cfgs, "Warning: model number {} not in cfgs dict!".format(model_name)
    convCfg = cfgs[model_name]

    model = ConvVGG(make_features(convCfg), **kwargs)
    return model
