# 导入minidom
from xml.dom import minidom

# 1.创建DOM树对象
dom = minidom.Document()
# 2.创建根节点。每次都要用DOM对象来创建任何节点。
root_node = dom.createElement('Order')
root_node.setAttribute('OrderID', '成都04-2019-5-321B-05')
root_node.setAttribute('Rmark', '主订单号')
# 3.用DOM对象添加根节点
dom.appendChild(root_node)

# 用DOM对象创建元素子节点
EOrder_node = dom.createElement('EOrder')
# 用父节点对象添加元素子节点
root_node.appendChild(EOrder_node)
# 设置该节点的属性
EOrder_node.setAttribute('EOrder', '成都04-2019-5-321B-05主卧衣柜')
EOrder_node.setAttribute('Rmark', '子订单号')


Code_node = dom.createElement('Code')
EOrder_node.appendChild(Code_node)
Code_node.setAttribute('Rmark', '板件编号(代号)')

CabinetName_node = dom.createElement('CabinetName')
Code_node.appendChild(CabinetName_node)
CabinetName_node.setAttribute('Rmark', '柜体名称')
name_text = dom.createTextNode('侧包顶')
CabinetName_node.appendChild(name_text)

CabinetType_node = dom.createElement('CabinetType')
Code_node.appendChild(CabinetType_node)
CabinetType_node.setAttribute('Rmark', '柜体型号')
name_text = dom.createTextNode('单元柜DYG-001')
CabinetType_node.appendChild(name_text)

BoardName_node = dom.createElement('BoardName')
Code_node.appendChild(BoardName_node)
BoardName_node.setAttribute('Rmark', '名称(板件)')
name_text = dom.createTextNode('左侧板')
BoardName_node.appendChild(name_text)



# 每一个结点对象（包括dom对象本身）都有输出XML内容的方法，如：toxml()--字符串, toprettyxml()--美化树形格式。

try:
    with open('dom_write.xml', 'w', encoding='UTF-8') as fh:
        # 4.writexml()第一个参数是目标文件对象，第二个参数是根节点的缩进格式，第三个参数是其他子节点的缩进格式，
        # 第四个参数制定了换行格式，第五个参数制定了xml内容的编码。
        dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
        print('写入xml OK!')
except Exception as err:
    print('错误信息：{0}'.format(err))