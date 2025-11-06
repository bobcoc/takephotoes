#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学生头像提取程序
使用 MediaPipe 检测人脸并提取头像区域
"""

import cv2
import mediapipe as mp
import os
from pathlib import Path


class HeadshotExtractor:
    """头像提取器"""
    
    def __init__(self, output_dir="cuted", scale_factor=1.8):
        """
        初始化头像提取器
        
        Args:
            output_dir: 输出目录名称
            scale_factor: 头像框扩展比例（相对于人脸检测框）
        """
        self.output_dir = output_dir
        self.scale_factor = scale_factor
        
        # 初始化 MediaPipe Face Detection
        self.mp_face_detection = mp.solutions.face_detection
        self.face_detection = self.mp_face_detection.FaceDetection(
            model_selection=1,  # 1表示全范围模型，适合距离较远的人脸
            min_detection_confidence=0.5
        )
        
    def extract_headshot(self, image_path, save_name=None):
        """
        从图片中提取头像
        
        Args:
            image_path: 输入图片路径
            save_name: 保存的文件名（不含扩展名），如果为None则使用原文件名
            
        Returns:
            bool: 是否成功提取
        """
        # 读取图片
        image = cv2.imread(str(image_path))
        if image is None:
            print(f"❌ 无法读取图片: {image_path}")
            return False
            
        # 转换为RGB（MediaPipe需要RGB格式）
        image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        
        # 检测人脸
        results = self.face_detection.process(image_rgb)
        
        if not results.detections:
            print(f"⚠️  未检测到人脸: {image_path}")
            return False
            
        # 获取图片尺寸
        h, w, _ = image.shape
        
        # 选择置信度最高的人脸（通常就是中央正面的人脸）
        best_detection = max(results.detections,
                             key=lambda d: d.score[0])
        
        # 获取人脸边界框
        bbox = best_detection.location_data.relative_bounding_box
        
        # 转换为像素坐标
        x = int(bbox.xmin * w)
        y = int(bbox.ymin * h)
        box_w = int(bbox.width * w)
        box_h = int(bbox.height * h)
        
        # 计算中心点
        center_x = x + box_w // 2
        center_y = y + box_h // 2
        
        # 扩展边界框以包含更多头部区域
        # 使用正方形框，以较大的边为基准
        box_size = max(box_w, box_h)
        expanded_size = int(box_size * self.scale_factor)
        
        # 计算新的边界框（正方形）
        new_x1 = max(0, center_x - expanded_size // 2)
        new_y1 = max(0, center_y - expanded_size // 2)
        new_x2 = min(w, center_x + expanded_size // 2)
        new_y2 = min(h, center_y + expanded_size // 2)
        
        # 裁剪头像区域
        headshot = image[new_y1:new_y2, new_x1:new_x2]
        
        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 确定保存的文件名
        if save_name is None:
            save_name = Path(image_path).stem
        
        # 保存头像（保持原格式）
        ext = Path(image_path).suffix
        output_path = os.path.join(self.output_dir, f"{save_name}{ext}")
        
        cv2.imwrite(output_path, headshot)
        print(f"✅ 成功提取头像: {output_path} (置信度: {best_detection.score[0]:.2f})")
        
        return True
    
    def batch_extract(self, input_dir=".", pattern="*.png"):
        """
        批量提取头像
        
        Args:
            input_dir: 输入目录
            pattern: 文件匹配模式（如 "*.png", "*.jpg" 等）
        """
        input_path = Path(input_dir)
        
        # 查找所有匹配的图片文件
        image_files = list(input_path.glob(pattern))
        
        if not image_files:
            print(f"⚠️  未找到匹配的图片文件: {pattern}")
            return
        
        print(f"📁 找到 {len(image_files)} 个图片文件")
        print(f"📂 输出目录: {self.output_dir}\n")
        
        success_count = 0
        failed_files = []
        
        for image_file in image_files:
            # 从文件名提取学号（去掉姓名部分）
            # 例如: "202510745_张殷瑞.png" -> "202510745"
            filename = image_file.stem
            if "_" in filename:
                student_id = filename.split("_")[0]
            else:
                student_id = filename
            
            # 提取头像
            if self.extract_headshot(image_file, student_id):
                success_count += 1
            else:
                failed_files.append(image_file.name)
        
        # 打印统计信息
        print(f"\n{'='*60}")
        print("✨ 处理完成！")
        print(f"   成功: {success_count}/{len(image_files)}")
        print(f"   失败: {len(failed_files)}/{len(image_files)}")
        
        if failed_files:
            print("\n❌ 失败的文件:")
            for filename in failed_files:
                print(f"   - {filename}")
    
    def __del__(self):
        """清理资源"""
        self.face_detection.close()


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="学生头像提取程序")
    parser.add_argument("-i", "--input", default=".",
                        help="输入目录（默认: 当前目录）")
    parser.add_argument("-o", "--output", default="cuted",
                        help="输出目录（默认: cuted）")
    parser.add_argument("-p", "--pattern", default="*.png",
                        help="文件匹配模式（默认: *.png）")
    parser.add_argument("-s", "--scale", type=float, default=1.8,
                        help="头像框扩展比例（默认: 1.8）")
    
    args = parser.parse_args()
    
    # 创建提取器
    extractor = HeadshotExtractor(
        output_dir=args.output,
        scale_factor=args.scale
    )
    
    # 批量处理
    extractor.batch_extract(
        input_dir=args.input,
        pattern=args.pattern
    )


if __name__ == "__main__":
    main()
