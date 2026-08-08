# -*- coding: utf-8 -*-
"""
core —— 与 UI / CLI 解耦的业务纯函数与数据模型。

Stage 0 目标：只抽取「不依赖 Photoshop COM、不依赖 Tk」的可测试纯函数，
不做任何行为变更。后续 Stage 再逐步把 Excel 数据层、图层索引、
Photoshop Session、renderer 迁入本包。
"""
