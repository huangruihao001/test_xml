from xml.dom import minidom
from ReadExcel import *

def creat_Order(OrderID, EOrderID):
    """
    创建主订单号和子订单号标签
    :param OrderID: 主订单号
    :param EOrderID: 子订单号
    :return:
    """
    Order_node.setAttribute('ID', OrderID)
    Order_node.setAttribute('Remark', '主订单号')
    # 3.用DOM对象添加根节点
    dom.appendChild(Order_node)

    Order_node.appendChild(EOrder_node)
    EOrder_node.setAttribute('ID', EOrderID)
    EOrder_node.setAttribute('Remark', '子订单号')

def cabinet_date(date_row, date_column, code_node):
    """
    生成板件详细信息的标签
    :param date_row:
    :param date_column:
    :param code_node:
    :return:
    """
    i_node = dom.createElement(str(excelOBJ.first_row[date_column].value))
    code_node.appendChild(i_node)
    name_text = dom.createTextNode(str(excelOBJ.date_row(date_row)[date_column].value))
    i_node.appendChild(name_text)

def board_code(date_row):
    """
    生成板件编号标签页
    :param date_row:
    :return:
    """
    code_node = dom.createElement('BoardCode')
    code_node.setAttribute('ID', str(excelOBJ.date_row(date_row)[4].value))
    code_node.setAttribute('Remark', '板件编号')
    EOrder_node.appendChild(code_node)

    # 生成板件详细信息的标签
    for i in range(1, excelOBJ.max_column):
        cabinet_date(2, i, code_node)


def board_code_all(date_row):
    """
    生成所有板件详细信息的标签
    :param date_row:
    :return:
    """
    for i in range(2, date_row+1):
        board_code(i)





def writeXML(Order):
    """
    按XML标准格式写入并保存
    :param Order: 订单号作为保存XML的文件名
    :return:
    """
    try:
        with open(Order+'.xml', 'w', encoding='UTF-8') as fh:
            # 4.writexml()第一个参数是目标文件对象，第二个参数是根节点的缩进格式，第三个参数是其他子节点的缩进格式，
            # 第四个参数制定了换行格式，第五个参数制定了xml内容的编码。
            dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
            print('写入xml OK!')
    except Exception as err:
        print('错误信息：{0}'.format(err))


if __name__ == '__main__':
    excelOBJ = ReadExcel("./圣萝莎erp输出输出参考.xlsx", "ERP表")
    dom = minidom.Document()# 创建根节点。每次都要用DOM对象来创建任何节点。
    Order_node = dom.createElement('Order')
    EOrder_node = dom.createElement('EOrder')
    creat_Order(excelOBJ.order, excelOBJ.Eorder)

    board_code_all(excelOBJ.max_row)




    writeXML(excelOBJ.order)
