import cv2
import numpy as np
import svgwrite

# 读取图像
image = cv2.imread(r"C:\Users\dell\Desktop\share\ComfyUI_windows_portable_nvidia_cu121_or_cpu\ComfyUI_windows_portable\ComfyUI\temp\ComfyUI_temp_onkup_00042_.png", cv2.IMREAD_GRAYSCALE)

# 二值化图像
_, binary_image = cv2.threshold(image, 127, 255, cv2.THRESH_BINARY)

# 寻找轮廓
contours, _ = cv2.findContours(binary_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# 选择最大的轮廓（假设是车身轮廓）
contour = max(contours, key=cv2.contourArea)

# 曲线拟合
epsilon = 0.01 * cv2.arcLength(contour, True)
curve_fitted = cv2.approxPolyDP(contour, epsilon, True)

# 转换拟合点到float32
curve_fitted = curve_fitted.astype(np.float32)

# 创建SVG文件
dwg = svgwrite.Drawing('output.svg', profile='tiny')

# 添加路径
path = dwg.path(d="M{} Z".format("L".join(["{} {},{}".format(pt[0], pt[1]) for pt in curve_fitted])))

# 将路径添加到SVG
dwg.add(path)

# 保存SVG文件
dwg.save()