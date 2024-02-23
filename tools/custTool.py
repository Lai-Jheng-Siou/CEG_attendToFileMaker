import configparser

def getConfig():
    config = configparser.ConfigParser()
    config.read('../config.ini', encoding='utf-8')
    return config