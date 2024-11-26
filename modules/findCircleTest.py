import cv2
import svgwrite

# 读取图像
image = cv2.imread(r"C:\Users\dell\Desktop\share\ComfyUI_windows_portable_nvidia_cu121_or_cpu"
                   r"\ComfyUI_windows_portable\ComfyUI\temp\ComfyUI_temp_hinkr_00011_.png", cv2.IMREAD_GRAYSCALE)

# 应用阈值来获取二值图像
_, thresh = cv2.threshold(image, 127, 255, cv2.THRESH_BINARY)

# 寻找轮廓
contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# 创建SVG文件
dwg = svgwrite.Drawing('output.svg', profile='tiny')

# 遍历轮廓，找到近似圆形并添加到SVG
for contour in contours:
    # 获取轮廓的近似圆形
    (x, y), radius = cv2.minEnclosingCircle(contour)

    # 四舍五入
    x, y, radius = int(x), int(y), int(radius)

    # 创建圆形并添加到SVG
    dwg.add(dwg.circle(center=(x, y), r=radius, fill='black'))

# 保存SVG文件
dwg.save()