"""
author:luojunwen
date:2021/2/3
"""
'''
创建文件夹
'''
import csv
import os

file1 = r"F:\PressureTestFiles\10picdoc"
file2 = r"F:\PressureTestFiles\10picdocx"
file3 = r"F:\PressureTestFiles\10picppt"
file4 = r"F:\PressureTestFiles\10picpptx"
file5 = r"F:\PressureTestFiles\10textdoc"
file6 = r"F:\PressureTestFiles\10textdocx"
file7 = r"F:\PressureTestFiles\10textppt"
file8 = r"F:\PressureTestFiles\10textpptx"
file9 = r"F:\PressureTestFiles\20picdoc"
file10 = r"F:\PressureTestFiles\20picdocx"
file11 = r"F:\PressureTestFiles\20picppt"
file12 = r"F:\PressureTestFiles\20picpptx"
file13 = r"F:\PressureTestFiles\20textdoc"
file14 = r"F:\PressureTestFiles\20textdocx"
file15 = r"F:\PressureTestFiles\20textppt"
file16 = r"F:\PressureTestFiles\20textpptx"
file17 = r"F:\PressureTestFiles\30picdoc"
file18 = r"F:\PressureTestFiles\30picdocx"
file19 = r"F:\PressureTestFiles\30picppt"
file20 = r"F:\PressureTestFiles\30picpptx"
file21 = r"F:\PressureTestFiles\30textdoc"
file22 = r"F:\PressureTestFiles\30textdocx"
file23 = r"F:\PressureTestFiles\30textppt"
file24 = r"F:\PressureTestFiles\30textpptx"
file25 = r"F:\PressureTestFiles\50picdoc"
file26 = r"F:\PressureTestFiles\50picdocx"
file27 = r"F:\PressureTestFiles\50picppt"
file28 = r"F:\PressureTestFiles\50picpptx"
file29 = r"F:\PressureTestFiles\50textdoc"
file30 = r"F:\PressureTestFiles\50textdocx"
file31 = r"F:\PressureTestFiles\50textppt"
file32 = r"F:\PressureTestFiles\50textpptx"
file33 = r"F:\PressureTestFiles\200x10xls"
file34 = r"F:\PressureTestFiles\200x10xlsx"


def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径


def dofile():
    mkdir(file1)
    mkdir(file2)
    mkdir(file3)
    mkdir(file4)
    mkdir(file5)
    mkdir(file6)
    mkdir(file7)
    mkdir(file8)
    mkdir(file9)
    mkdir(file10)
    mkdir(file11)
    mkdir(file12)
    mkdir(file13)
    mkdir(file14)
    mkdir(file15)
    mkdir(file16)
    mkdir(file17)
    mkdir(file18)
    mkdir(file19)
    mkdir(file20)
    mkdir(file21)
    mkdir(file22)
    mkdir(file23)
    mkdir(file24)
    mkdir(file25)
    mkdir(file26)
    mkdir(file27)
    mkdir(file28)
    mkdir(file29)
    mkdir(file30)
    mkdir(file31)
    mkdir(file32)
    mkdir(file33)
    mkdir(file34)



