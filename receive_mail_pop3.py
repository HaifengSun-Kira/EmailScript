#encoding:utf8
import poplib
import time
import re
import configparser
from email.parser import Parser
from email.header import decode_header
from email.header import Header

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def get_att(msg, dir = ""):
    # for header in ['From', 'To', 'Subject']:
    #     value = msg.get(header, '')
    #     if value:
    #         if header=='Subject':
    #             value = decode_str(value)
    #         else:
    #             hdr, addr = parseaddr(value)
    #             name = decode_str(hdr)
    #             value = u'%s <%s>' % (name, addr)
    #     print('%s: %s' % (header, value))
    for part in msg.walk():
        file_name = part.get_filename()  # 获取附件名称类型
        if file_name:
            h = Header(file_name)
            dh = decode_header(h)  # 对附件名称进行解码
            filename = dh[0][0]
            if dh[0][1]:
                filename = decode_str(str(filename, dh[0][1]))  # 将附件名称可读化
                print(filename)
                # filename = filename.encode("utf-8")
            data = part.get_payload(decode=True)  # 下载附件
            att_file = open(dir + filename, 'wb')  # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
            att_file.write(data)  # 保存附件
            att_file.close()


if __name__ == "__main__":
    starttime = "20180315"
    endtime = "20180322"
    title = r"chapter2_homework_\d{9}"
    att_dir = ""

    pattern = re.compile(title)

    conf = configparser.ConfigParser()
    conf.read("email.cfg")

    email = conf.get("email", "email")
    password = conf.get("email", "password")
    pop3_server = conf.get("email", "pop3")

    # 连接到POP3服务器:
    server = poplib.POP3(pop3_server)
    # 可以打开或关闭调试信息:
    server.set_debuglevel(0)
    # 可选:打印POP3服务器的欢迎文字:
    print(server.getwelcome().decode('utf-8'))

    # 身份认证:
    server.user(email)
    server.pass_(password)

    # stat()返回邮件数量和占用空间:
    print('Messages: %s. Size: %s' % server.stat())
    # list()返回所有邮件的编号:
    resp, mails, octets = server.list()
    # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
    # print(mails)

    index = len(mails)

    for i in range(index, 0, -1):
        # 倒序遍历邮件
        resp, lines, octets = server.retr(i)
        # lines存储了邮件的原始文本的每一行,
        # 邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 解析邮件:
        msg = Parser().parsestr(msg_content)
        # 获取邮件主题
        subject = decode_str(msg.get("Subject", ''))
        if not pattern.match(subject):
            continue
        # 获取邮件时间
        date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')  # 格式化收件时间
        date2 = time.strftime("%Y%m%d", date1)  # 邮件时间格式转换
       # print(date2)
        if (date2 < starttime) | (date2 > endtime):
            break
        f_list = get_att(msg, att_dir)  # 获取附件

    # 可以根据邮件索引号直接从服务器删除邮件:
    # server.dele(index)
    # 关闭连接:
    server.quit()
