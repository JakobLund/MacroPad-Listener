from datetime import datetime


def write_to_log(msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg_ = f"{timestamp} - {msg}"
    log_file = open("com_port_listener.log", "a")
    log_file.write(msg_ + "\n")
    log_file.close()
