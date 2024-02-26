import logging  #import the logging module

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

def setupLogger(name, log_file, level=logging.INFO):
    """To setup as many loggers as you want"""

    handler = logging.FileHandler(log_file)        
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    
    return logger
   