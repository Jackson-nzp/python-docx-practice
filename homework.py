#安装包 pip install python-docx 
#国内镜像站： -i https://pypi.tuna.tsinghua.edu.cn/simple

from docx.document import Document
from docx.text.paragraph import Paragraph #段落处理
from docx.shared import RGBColor  #设置颜色
from docx import Document
from docx.shared import Pt  #设置字号，间距
from docx.oxml.ns import qn  # 设置中文字体

document = Document()#新建文档
number = document.add_paragraph().add_run('200825027') #添加段落及段内文字
number.font.size= Pt(26) #设置文字字号
number.font.color.rgb=RGBColor(128, 0, 128) #设置文字颜色，紫色rgb
#下同
college = document.add_paragraph().add_run('商学院')
college.font.size = Pt(26)
college.font.color.rgb=RGBColor(128, 0, 128)

paragraph_3= document.add_paragraph()
name=paragraph_3.add_run('王柯蓉')
name.font.size = Pt(26)
name.font.color.rgb=RGBColor(128, 0, 128)

paragraph_format = paragraph_3.paragraph_format #仅能处理段落，故paragraph_3不直接add，先设置个段落出来
paragraph_format.space_after = Pt(45) #60像素以windows默认dpi：96转换为45pt，详见：https://www.cnblogs.com/zhenzhong/p/3348567.html

date=document.add_paragraph()
year=date.add_run('2022年')
year.font.size = Pt(26)
year.font.color.rgb=RGBColor(255, 0, 0)

month=date.add_run('11月')
month.font.size = Pt(26)
month.font.color.rgb=RGBColor(0,255,0)

day=date.add_run('15日')
day.font.size = Pt(26)
day.font.color.rgb=RGBColor(0, 0, 255)

document.save('homework.docx') #文档保存