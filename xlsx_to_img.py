import os
from spire.xls import *
from spire.xls.common import *

def excel_to_image(input_dir, output_dir):
    """
    遍历指定目录下的所有Excel文件，将指定单元格区域保存为图片

    参数:
    input_dir: 包含Excel文件的输入目录
    output_dir: 保存截图图片的输出目录
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_dir):
        if filename.endswith(('.xlsx', '.xls')):
            try:
                workbook = Workbook()

                file_path = os.path.join(input_dir, filename)

                workbook.LoadFromFile(file_path)

                # 获取第一个工作表
                sheet = workbook.Worksheets[0]

                # 将指定单元格区域转换为图片, 参数: 起始行, 起始列, 结束行, 结束列
                image = sheet.ToImage(8, 11, 44, 20)

                # 构建输出图片路径 (保持原文件名)
                image_name = os.path.splitext(filename)[0] + ".png"
                image_path = os.path.join(output_dir, image_name)

                image.Save(image_path)
                print(f"已保存: {image_path}")

            except Exception as e:
                print(f"处理文件 {filename} 时出错: {str(e)}")
            finally:
                workbook.Dispose()

if __name__ == "__main__":
    # Excel文件所在目录
    input_directory = "xlsx"
    # 截图保存目录
    output_directory = "picture"

    # 执行转换
    excel_to_image(input_directory, output_directory)
    print("所有文件处理完成！")
