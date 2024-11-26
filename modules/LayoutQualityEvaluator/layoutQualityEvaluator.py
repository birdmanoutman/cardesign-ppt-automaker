import os
import re

import cv2
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
from matplotlib.font_manager import FontProperties

# 设置中文字体
font = FontProperties(fname=r'C:\Windows\Fonts\msyh.ttc')  # 这里选择一个系统中存在的中文字体文件


def calculate_overall_score(contrast, alignment, white_space, color_consistency, visual_hierarchy):
    contrast_score = min(contrast / 50.0, 1.0) * 2.0
    alignment_score = min(alignment / 20.0, 1.0) * 2.0
    white_space_score = min(white_space / 0.5, 1.0) * 2.0
    color_consistency_score = min(color_consistency, 1.0) * 2.0
    visual_hierarchy_score = min(visual_hierarchy / 10.0, 1.0) * 2.0
    total_score = contrast_score + alignment_score + white_space_score + color_consistency_score + visual_hierarchy_score
    return contrast_score, alignment_score, white_space_score, color_consistency_score, visual_hierarchy_score, total_score


def plot_scores(scores):
    # 按照幻灯片名称进行自然排序
    sorted_scores = dict(sorted(scores.items(), key=lambda item: natural_keys(item[0])))

    labels = list(sorted_scores.keys())
    contrast_scores = [score[0] for score in sorted_scores.values()]
    alignment_scores = [score[1] for score in sorted_scores.values()]
    white_space_scores = [score[2] for score in sorted_scores.values()]
    color_consistency_scores = [score[3] for score in sorted_scores.values()]
    visual_hierarchy_scores = [score[4] for score in sorted_scores.values()]

    fig, ax = plt.subplots(figsize=(10, 10))

    ax.barh(labels, contrast_scores, color='blue', label='Contrast')
    ax.barh(labels, alignment_scores, left=contrast_scores, color='green', label='Alignment')
    ax.barh(labels, white_space_scores, left=np.array(contrast_scores) + np.array(alignment_scores), color='red',
            label='White Space')
    ax.barh(labels, color_consistency_scores,
            left=np.array(contrast_scores) + np.array(alignment_scores) + np.array(white_space_scores), color='cyan',
            label='Color Consistency')
    ax.barh(labels, visual_hierarchy_scores,
            left=np.array(contrast_scores) + np.array(alignment_scores) + np.array(white_space_scores) + np.array(
                color_consistency_scores), color='magenta', label='Visual Hierarchy')

    ax.set_xlabel('Score', fontproperties=font)
    ax.set_ylabel('Image', fontproperties=font)
    ax.set_title('Layout Quality Scores', fontproperties=font)
    ax.legend()
    plt.xticks(fontproperties=font)
    plt.yticks(fontproperties=font)
    plt.show()


# 自然排序所需的辅助函数
def atoi(text):
    return int(text) if text.isdigit() else text


def natural_keys(text):
    return [atoi(c) for c in re.split(r'(\d+)', text)]


# 计算对比度
def calculate_contrast(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    contrast = gray.std()
    return contrast


# 计算对齐
def calculate_alignment(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=50, minLineLength=50, maxLineGap=10)
    alignment_score = 0
    if lines is not None:
        alignment_score = len(lines)
    return alignment_score


# 计算空白
def calculate_white_space(image):
    # 将图像转换为灰度图像
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # 应用高斯模糊
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

    # 使用自适应阈值法进行二值化处理
    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 11, 2)

    # 使用形态学操作清理图像
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    clean = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    clean = cv2.morphologyEx(clean, cv2.MORPH_OPEN, kernel)

    # 计算背景比例
    background_ratio = np.sum(clean == 0) / (clean.shape[0] * clean.shape[1])

    return background_ratio


# 计算颜色一致性
import cv2
import numpy as np


def calculate_color_consistency(image):
    # 将图像从BGR颜色空间转换为LAB颜色空间
    lab_image = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)

    # 计算全局平均颜色
    mean_color = np.mean(lab_image, axis=(0, 1))

    # 计算每个像素颜色与全局平均颜色的欧几里得距离
    distance = np.linalg.norm(lab_image - mean_color, axis=2)

    # 计算距离的平均值作为颜色一致性的反向指标
    color_consistency_distance = 1.0 / (1.0 + np.mean(distance))

    # 计算颜色直方图的均匀性
    hist_a = cv2.calcHist([lab_image], [1], None, [256], [0, 256])
    hist_b = cv2.calcHist([lab_image], [2], None, [256], [0, 256])

    hist_a = hist_a / np.sum(hist_a)
    hist_b = hist_b / np.sum(hist_b)

    uniformity_a = -np.sum(hist_a * np.log(hist_a + 1e-5))
    uniformity_b = -np.sum(hist_b * np.log(hist_b + 1e-5))

    color_consistency_hist = 1.0 / (1.0 + uniformity_a + uniformity_b)

    # 综合两种方法的结果
    color_consistency = (color_consistency_distance + color_consistency_hist) / 2.0

    return color_consistency


# 计算视觉层次
def calculate_visual_hierarchy(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    hierarchy_score = len(contours)
    return hierarchy_score


# 处理单个图像文件
def process_single_image(image_path):
    try:
        pil_image = Image.open(image_path)
        image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

        contrast = calculate_contrast(image)
        alignment = calculate_alignment(image)
        white_space = calculate_white_space(image)
        color_consistency = calculate_color_consistency(image)
        visual_hierarchy = calculate_visual_hierarchy(image)
        scores = calculate_overall_score(contrast, alignment, white_space, color_consistency, visual_hierarchy)

        print(f"Image: {image_path}, Scores: {scores}")
        return scores
    except Exception as e:
        print(f"Error: Couldn't read image {image_path}. Exception: {str(e)}")
        return None


# 批量处理图像目录中的所有图像
def process_images(image_dir):
    scores = {}
    for image_name in os.listdir(image_dir):
        image_path = os.path.join(image_dir, image_name)
        if not image_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
            continue
        score = process_single_image(image_path)
        if score is not None:
            scores[image_name] = score
    return scores


import csv


def save_scores_to_csv(scores, file_path):
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Image', 'Contrast', 'Alignment', 'White Space', 'Color Consistency', 'Visual Hierarchy', 'Total Score'])
        for image, score in scores.items():
            writer.writerow([image] + list(score))


# 主函数
if __name__ == "__main__":
    image_dir = r"C:\Users\dell\Desktop\230305_百人会演讲_顺序调整2zez"  # 更改为你的图像目录
    scores = process_images(image_dir)

    # 保存评分结果到CSV文件
    save_scores_to_csv(scores, 'layout_quality_scores.csv')

    # 可视化评分结果
    plot_scores(scores)
