#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
摄像头检测工具
检测系统中可用的摄像头设备
"""

import cv2

def detect_cameras():
    """检测可用的摄像头"""
    available_cameras = []
    
    print("正在检测可用的摄像头...")
    
    # 检测前10个可能的摄像头索引
    for index in range(10):
        cap = cv2.VideoCapture(index)
        if cap.isOpened():
            # 尝试读取一帧来确认摄像头真的可用
            ret, frame = cap.read()
            if ret:
                # 获取摄像头信息
                width = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
                height = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
                fps = cap.get(cv2.CAP_PROP_FPS)
                
                camera_info = {
                    'index': index,
                    'width': int(width),
                    'height': int(height),
                    'fps': fps
                }
                available_cameras.append(camera_info)
                print(f"✅ 摄像头 {index}: {int(width)}x{int(height)} @ {fps:.1f}fps")
            cap.release()
        else:
            # 如果连续3个索引都没有摄像头，就停止检测
            if index > 2 and len(available_cameras) == 0:
                break
    
    if not available_cameras:
        print("❌ 未检测到任何可用的摄像头")
    else:
        print(f"\n共检测到 {len(available_cameras)} 个可用摄像头")
    
    return available_cameras

if __name__ == "__main__":
    cameras = detect_cameras()
    
    if cameras:
        print("\n摄像头详细信息:")
        for camera in cameras:
            print(f"索引 {camera['index']}: {camera['width']}x{camera['height']} @ {camera['fps']:.1f}fps")