import time

import re
import socket
import subprocess
import xml.etree.ElementTree as ET


def run_julius():
    """
    Juliusの起動.
    """
    print('start Julius')
    subprocess.Popen('/home/pi/speech/shell/run_julius.sh', shell=True)
    time.sleep(5)
    print('ready!')


def kill_julius():
    """
    Juliusの終了.
    """
    print('finish Julius')
    subprocess.Popen('/home/pi/speech/shell/kill_julius.sh', shell=True)


def extract_words(response):
    """
    Julius音声認識結果からテキストを抽出.
    :param response: socket.recv()の結果(文字列)
    :return: 音声認識結果文字列
    """
    # 正規表現で認識結果部分を切り出し
    xml_recogout = re.search(
        r'<RECOGOUT>.+</RECOGOUT>',
        response,
        flags=re.DOTALL)

    if xml_recogout is None:
        return

    # XMLパース
    recogout = ET.fromstring(xml_recogout[0])
    words = []
    for whypo in recogout[0]:
        attrib = whypo.attrib
        if attrib['WORD'] not in ('<s>', '</s>'):
            words.append(attrib['WORD'])

    return ''.join(words)


def julius_speech_to_text(callback=None):
    """
    Julius音声認識開始.
    :param callback: コールバック
    """
    host = '127.0.0.1'
    port = 10500
    client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client.connect((host, port))
    while True:
        time.sleep(0.1)
        response = client.recv(4096).decode('utf-8')

        text = extract_words(response)
        if text is None:
            continue

        text = text + '\n' 
        print(text)
        host1 = '192.168.11.53'
        port1 = 7000
        sock1 = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        message = text.encode('utf-8')
        sock1.sendto(message, (host1, port1))

        if callback is not None:
            callback(text)


if __name__ == '__main__':
    run_julius()
    try:
        julius_speech_to_text()

    except KeyboardInterrupt:
        print('keyboard interrupt')

    finally:
        kill_julius()
