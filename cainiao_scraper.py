from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

class CainiaoScraper:
    def __init__(self):
        try:
            # 設定 Chrome 選項
            chrome_options = Options()
            
            # 添加新的參數來解決警告
            chrome_options.add_argument('--disable-gpu')  # 禁用 GPU 硬件加速
            chrome_options.add_argument('--no-sandbox')  # 禁用沙盒模式
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-web-security')  # 禁用網頁安全性檢查
            chrome_options.add_argument('--disable-webgl')  # 禁用 WebGL
            chrome_options.add_argument('--disable-notifications')  # 禁用通知
            chrome_options.add_argument('--disable-logging')  # 禁用日誌
            chrome_options.add_argument('--log-level=3')  # 只顯示重要訊息
            chrome_options.add_argument('--ignore-certificate-errors')  # 忽略證書錯誤
            chrome_options.add_argument('--disable-software-rasterizer')  # 禁用軟體光柵化器
            
            # 添加實驗性功能參數
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
            
            # 使用 webdriver_manager 自動下載對應版本的 ChromeDriver
            print("正在初始化 Chrome WebDriver...")
            self.driver = webdriver.Chrome(
                service=ChromeService(ChromeDriverManager().install()),
                options=chrome_options
            )
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 20)
            print("Chrome WebDriver 初始化成功！")
            
        except Exception as e:
            print(f"初始化失敗：{str(e)}")
            raise
        
    def login(self, username, password):
        try:
            print("正在訪問菜鳥網站...")
            self.driver.get("https://desk.cainiao.com/unified/ticketManage/processingTicketManage")
            
            # 等待登入頁面加載
            time.sleep(5)
            print("頁面加載完成")
            
            # 這裡需要加入實際的登入邏輯
            # 例如：
            # username_input = self.wait.until(
            #     EC.presence_of_element_located((By.ID, "username"))
            # )
            # password_input = self.wait.until(
            #     EC.presence_of_element_located((By.ID, "password"))
            # )
            # username_input.send_keys(username)
            # password_input.send_keys(password)
            # login_button = self.wait.until(
            #     EC.element_to_be_clickable((By.ID, "submit"))
            # )
            # login_button.click()
            
            print("等待手動登入...")
            input("請在瀏覽器中完成登入後，按下 Enter 繼續...")
            
        except Exception as e:
            print(f"登入過程發生錯誤：{str(e)}")
            raise
        
    def get_ticket_info(self):
        try:
            print("等待頁面完全載入...")
            time.sleep(5)  # 等待頁面完全載入
            
            print("正在尋找處理中的工單按鈕...")
            # 使用更多的定位方式嘗試找到按鈕
            try:
                # 方法1：使用文字內容定位
                processing_btn = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '处理中的工单')]"))
                )
            except:
                try:
                    # 方法2：使用class定位
                    processing_btn = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".processing-ticket-btn"))
                    )
                except:
                    # 方法3：使用完整的XPath（請根據實際網頁結構調整）
                    processing_btn = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@class='ticket-list']//button[contains(@class, 'processing')]"))
                    )
            
            processing_btn.click()
            print("已點擊處理中的工單按鈕")
            
            # 等待工單列表加載
            time.sleep(5)
            
            tickets_data = []
            print("正在獲取工單列表...")
            
            # 嘗試不同的定位方式獲取工單列表
            try:
                # 方法1：使用表格定位
                ticket_links = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table tbody tr td:first-child a"))
                )
            except:
                try:
                    # 方法2：使用列表定位
                    ticket_links = self.wait.until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ticket-list .ticket-item .ticket-number"))
                    )
                except:
                    # 方法3：使用更寬鬆的定位方式
                    ticket_links = self.wait.until(
                        EC.presence_of_all_elements_located((By.XPATH, "//*[contains(@class, 'ticket-number') or contains(@class, 'order-number')]"))
                    )
            
            print(f"找到 {len(ticket_links)} 個工單")
            
            for index, ticket in enumerate(ticket_links, 1):
                try:
                    print(f"正在處理第 {index} 個工單...")
                    ticket_number = ticket.text
                    print(f"工單號：{ticket_number}")
                    
                    # 使用 JavaScript 點擊元素，避免元素不可點擊的問題
                    self.driver.execute_script("arguments[0].click();", ticket)
                    time.sleep(3)
                    
                    print("正在獲取運單號...")
                    try:
                        # 方法1：使用標準定位
                        tracking_number = self.wait.until(
                            EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'运单号') or contains(text(),'运单号')]/following-sibling::*"))
                        ).text
                    except:
                        # 方法2：使用更寬鬆的定位
                        tracking_number = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, ".tracking-number, .waybill-number"))
                        ).text
                    
                    print("正在獲取工單描述...")
                    try:
                        # 方法1：使用標準定位
                        ticket_description = self.wait.until(
                            EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'工单描述') or contains(text(),'工单描述')]/following-sibling::*"))
                        ).text
                    except:
                        # 方法2：使用更寬鬆的定位
                        ticket_description = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, ".ticket-description, .description"))
                        ).text
                    
                    tickets_data.append({
                        '工单号': ticket_number,
                        '运单号': tracking_number,
                        '工单描述': ticket_description
                    })
                    
                    print(f"成功獲取工單資料：{ticket_number}")
                    
                    # 使用 JavaScript 返回上一頁
                    self.driver.execute_script("window.history.go(-1)")
                    time.sleep(3)
                    
                except Exception as e:
                    print(f"處理工單 {ticket_number} 時發生錯誤：{str(e)}")
                    # 嘗試返回列表頁
                    try:
                        self.driver.execute_script("window.history.go(-1)")
                    except:
                        pass
                    time.sleep(3)
                    continue
            
            if tickets_data:
                print("正在匯出資料到 Excel...")
                df = pd.DataFrame(tickets_data)
                df.to_excel('工單資料.xlsx', index=False, encoding='utf-8-sig')
                print("資料已成功匯出到 Excel")
            else:
                print("沒有找到任何工單資料")
            
            return tickets_data
            
        except TimeoutException as e:
            print(f"發生超時錯誤: {str(e)}")
            return None
        except Exception as e:
            print(f"發生錯誤: {str(e)}")
            return None
            
    def close(self):
        print("正在關閉瀏覽器...")
        self.driver.quit()
        print("瀏覽器已關閉")

def main():
    scraper = None
    try:
        print("開始執行爬蟲程式...")
        scraper = CainiaoScraper()
        
        # 這裡需要填入實際的帳號密碼
        scraper.login('your_username', 'your_password')
        ticket_data = scraper.get_ticket_info()
        
        if ticket_data:
            print("資料擷取成功！")
            print(f"共擷取 {len(ticket_data)} 筆工單資料")
        else:
            print("資料擷取失敗！")
            
    except Exception as e:
        print(f"程式執行過程中發生錯誤：{str(e)}")
    finally:
        if scraper:
            scraper.close()

if __name__ == "__main__":
    main() 