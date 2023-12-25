from selenium import webdriver
from selenium.webdriver.common.by import By
import time

def run_webdriver():
    # 配置Chrome浏览器
    options = webdriver.ChromeOptions()
    options.add_experimental_option('detach', True)
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(options=options)

    try:
        # 打开目标网站
        driver.get('https://novelai.net/image')

        count = 1
        clicks_per_generation = 10

        # 记录开始时间
        start_time = time.time()
        last_print_time = start_time

        while True:
            try:
                # 尝试点击按钮
                driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div[4]/div[1]/div[5]/button').click()

                if count % clicks_per_generation == 0:
                    print(f"生成点击次数: {count}")

                time.sleep(10)

                count += 1

                # 获取当前时间
                current_time = time.time()

                # 如果已经过了一分钟，打印已运行时间
                if current_time - last_print_time >= 60:
                    minutes_elapsed = int((current_time - start_time) // 60)
                    seconds_elapsed = int((current_time - start_time) % 60)
                    print(f"已运行时间: {minutes_elapsed} 分钟 {seconds_elapsed} 秒")
                    last_print_time = current_time  # 更新上次打印时间

            except Exception as e:
                # 没找到生成按钮时，继续循环，不输出报错信息
                pass

    except KeyboardInterrupt:
        pass  # 捕获键盘中断信号，例如Ctrl+C

if __name__ == '__main__':
    run_webdriver()
