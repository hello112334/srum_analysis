import multiprocessing
import time

def worker():
    while True:
        pass

if __name__ == '__main__':
    # 获取CPU核心数
    num_cores = multiprocessing.cpu_count()
    print(f'Number of cores: {num_cores}')

    # 开始多进程运行
    processes = []
    for i in range(num_cores):
        p = multiprocessing.Process(target=worker)
        p.start()
        processes.append(p)

    # 运行5秒后停止所有进程
    time.sleep(5)
    for p in processes:
        p.terminate()