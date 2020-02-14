from ConfigParser import ConfigParser

def getpath(type1,type2):
    config = ConfigParser()
    config.read('config.ini')
    result = config.get(type1,type2)
    return result