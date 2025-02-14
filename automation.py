import json
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import traceback
import re
from datetime import datetime, timedelta
import os
import pyperclip
from selenium.webdriver.common.keys import Keys

def show_notification(driver, message):
    """顯示瀏覽器通知"""
    try:
        # 轉義特殊字符
        message = message.replace("'", "\\'").replace("\n", "\\n")
        script = (
            "var notification = document.createElement('div');"
            f"notification.textContent = '{message}';"
            "notification.style.position = 'fixed';"
            "notification.style.top = '20px';"
            "notification.style.right = '20px';"
            "notification.style.backgroundColor = '#4CAF50';"
            "notification.style.color = 'white';"
            "notification.style.padding = '15px';"
            "notification.style.borderRadius = '5px';"
            "notification.style.zIndex = '9999';"
            "notification.style.maxWidth = '300px';"
            "notification.style.boxShadow = '0 4px 8px rgba(0,0,0,0.1)';"
            "document.body.appendChild(notification);"
            "setTimeout(function() {"
            "    notification.style.transition = 'opacity 0.5s';"
            "    notification.style.opacity = '0';"
            "    setTimeout(function() {"
            "        notification.remove();"
            "    }, 500);"
            "}, 3000);"
        )
        driver.execute_script(script)
    except Exception as e:
        print(f"顯示通知時發生錯誤: {str(e)}")

def log_message(driver, message):
    """同時執行 print 和瀏覽器通知"""
    print(message)
    if driver:
        show_notification(driver, message)

driver = None
wait = None

def setup_driver():
    """設置並返回 Chrome WebDriver"""
    global driver, wait  # 添加全局變量聲明
    chrome_options = Options()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--disable-infobars')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_experimental_option("prefs", {
        "profile.password_manager_enabled": False,
        "credentials_enable_service": False,
    })
    
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=chrome_options
    )
    wait = WebDriverWait(driver, 30)  # 增加默認等待時間
    print("瀏覽器已啟動")
    log_message(driver, "瀏覽器已啟動")
    return driver

def retry_operation(operation, max_retries=5, delay=2):
        for i in range(max_retries):
            try:
                return operation()
            except Exception as e:
                if i == max_retries - 1:
                    raise e
                print(f"操作失敗,重試中... ({i+1}/{max_retries})")
                log_message(driver, f"操作失敗,重試中... ({i+1}/{max_retries})")
                time.sleep(delay)

def login_to_cainiao(driver):
    """登入菜鳥系統"""
    try:
        # 訪問初始頁面
        driver.get("https://desk.cainiao.com")
        log_message(driver, "正在訪問菜鳥系統...")
        
        # 等待並切換到登入 iframe
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )
        driver.switch_to.frame(iframe)
        log_message(driver, "已切換到登入框架")
        
        # 使用 driver 創建新的 WebDriverWait 實例
        wait = WebDriverWait(driver, 30)
        
        username_input = wait.until(
            EC.presence_of_element_located((By.NAME, 'fm-login-id'))
        )
        username_input.clear()
        username_input.send_keys('Abby.Chen@presco.ws')
        print("已輸入用戶名")

        password_input = wait.until(
            EC.presence_of_element_located((By.NAME, 'fm-login-password'))
        )
        password_input.clear()
        password_input.send_keys('Abby1023')
        print("已輸入密碼")

        login_button = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'fm-button'))
        )
        login_button.click()
        print("已點擊登入按鈕")
        
        # 切換回主框架
        driver.switch_to.default_content()
        
        # 等待登入成功
        time.sleep(5)
        
        # 直接導向到工單頁面
        driver.get("https://desk.cainiao.com/unified/myTask/pendingTask")
        log_message(driver, "正在導向工單頁面...")
        
        # 等待頁面加載完成
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        
        # 驗證是否成功進入工單頁面
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            log_message(driver, "登入成功並已進入工單頁面")
            return True
        except:
            log_message(driver, "無法進入工單頁面，可能登入失敗")
            return False
        
    except Exception as e:
        log_message(driver, f"登入失敗: {str(e)}")
        return False

def get_tracking_info(tracking_number):
    """獲取運單追蹤信息"""
    token_url = "https://ecapi.sp88.tw/api/Token"
    token_headers = {
        "Content-Type": "application/json",
    }
    token_data = {
        "account": "ESCS",
        "password": "SDG3jdkd59@1"
    }
    token_data_str = json.dumps(token_data)
    token = None
    token_success = False

    try:
        token_response = requests.post(token_url, headers=token_headers, data=token_data_str, timeout=30)
        
        if token_response.status_code == 200:
            token_data = token_response.json()
            if token_data.get("token") != None:
                token = token_data.get("token")
                token_success = True
            else:
                error_msg = token_data.get("Message", "未知錯誤")
                print(f"獲取 token 失敗,API 返回錯誤: {error_msg}")
                return None, "獲取 token 失敗"
        else:
            print(f"HTTP 請求失敗,狀態碼: {token_response.status_code}")
            return None, "HTTP 請求失敗"

        if token_success:
            track_url = f"https://ecapi.sp88.tw/api/track/B2C?eshopId=74A&shipmentNo={tracking_number}"
            track_headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}"
            }
            
            track_response = requests.get(track_url, headers=track_headers)
            
            if track_response.status_code == 200:
                track_data = track_response.json()
                return track_data, None
            else:
                return None, f"追蹤請求失敗,狀態碼: {track_response.status_code}"
                
    except requests.exceptions.Timeout:
        return None, "請求超時"
    except requests.exceptions.RequestException as e:
        return None, f"請求異常: {str(e)}"
    except json.JSONDecodeError as e:
        return None, f"解析 JSON 響應失敗: {str(e)}"

def process_order(driver, order_link):
    """處理單個工單"""
    original_window = driver.current_window_handle
    try:
        # 創建新的 WebDriverWait 實例
        wait = WebDriverWait(driver, 30)
        
        # 獲取當前工單號和URL
        current_order_id = order_link.text.strip()
        current_order_url = order_link.get_attribute('href')
        
        # 點擊工單連結，在新視窗中打開
        driver.execute_script("window.open(arguments[0], '_blank');", current_order_url)
        
        # 等待新視窗打開並切換到新視窗
        wait.until(lambda d: len(d.window_handles) > 1)
        new_window = [handle for handle in driver.window_handles if handle != original_window][0]
        driver.switch_to.window(new_window)
        
        # 等待頁面加載完成
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(3)  # 等待頁面元素完全加載
        
        try:
            # 尋找業務單據號（物流單號）
            tracking_number_element = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "td[data-next-table-col='4']"))
            )
            tracking_number = tracking_number_element.text.strip()
            
            # 獲取追蹤信息
            tracking_info, error = get_tracking_info(tracking_number)
            if error:
                print(f"獲取追蹤信息失敗: {error}")
                return
            
            # 點擊留言按鈕
            comment_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., '留言')]"))
            )
            comment_button.click()
            time.sleep(2)
            
            # 等待文本框出現並輸入追蹤信息
            textarea = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "textarea.next-input"))
            )
            textarea.clear()
            textarea.send_keys(str(tracking_info))
            
            # 點擊提交按鈕
            submit_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., '提交')]"))
            )
            submit_button.click()
            time.sleep(2)
            
            # 點擊結單按鈕
            close_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., '結單')]"))
            )
            close_button.click()
            time.sleep(2)
            
            # 保存處理記錄
            save_processed_order(current_order_id, current_order_url)
            log_message(driver, f"工單 {current_order_id} 處理完成")
            
        except Exception as e:
            print(f"處理工單內容時發生錯誤: {str(e)}")
            
    except Exception as e:
        print(f"處理工單時發生錯誤: {str(e)}")
        
    finally:
        # 關閉新視窗並切換回原始視窗
        if driver.current_window_handle != original_window:
            driver.close()
            driver.switch_to.window(original_window)
            time.sleep(1)  # 等待視窗切換完成

def load_processed_orders():
    """載入已處理的工單記錄"""
    try:
        if os.path.exists('processed_orders.json'):
            with open('processed_orders.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"讀取處理記錄時發生錯誤: {str(e)}")
    return {}

def save_processed_order(order_id, order_url):
    """保存已處理的工單記錄"""
    try:
        processed_orders = load_processed_orders()
        if order_id not in processed_orders:
            processed_orders[order_id] = {
                'url': order_url,
                'processed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open('processed_orders.json', 'w', encoding='utf-8') as f:
                json.dump(processed_orders, f, ensure_ascii=False, indent=2)
            print(f"已將工單 {order_id} 添加到處理記錄")
    except Exception as e:
        print(f"保存處理記錄時發生錯誤: {str(e)}")

def get_unique_order_links(driver, wait):
    """獲取唯一的工單連結"""
    try:
        # 等待頁面加載完成
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(3)  # 等待動態內容加載
        
        # 使用更精確的選擇器，定位工單號列的連結
        order_links = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td[data-next-table-col='1'] a"))
        )
        
        if not order_links:
            print("未找到工單連結")
            return []
            
        unique_order_links = []
        seen_order_ids = set()
        
        for link in order_links:
            order_id = link.text.strip()
            if order_id and order_id not in seen_order_ids:
                seen_order_ids.add(order_id)
                unique_order_links.append(link)
        
        print(f"找到 {len(unique_order_links)} 個唯一工單")
        return unique_order_links
        
    except Exception as e:
        print(f"獲取工單連結時發生錯誤: {str(e)}")
        return []

def main():
    """主程序"""
    driver = None
    try:
        driver = setup_driver()
        if not login_to_cainiao(driver):
            return
            
        while True:
            try:
                # 訪問待處理工單頁面
                # driver.get("https://desk.cainiao.com/unified/myTask/pendingTask")
                # log_message(driver, "正在訪問待處理工單頁面...")
                # time.sleep(5)  # 等待頁面初始加載
                
                # 等待頁面完全加載
                # wait = WebDriverWait(driver, 30)
                # wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
                
                # 等待表格元素出現
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".next-table-body")))
                    log_message(driver, "表格已加載")
                except:
                    log_message(driver, "找不到工單表格，重新整理頁面...")
                    driver.refresh()
                    time.sleep(5)
                    continue
                
                # 獲取工單連結
                try:
                    order_links = get_unique_order_links(driver, wait)
                    if not order_links:
                        log_message(driver, "當前無待處理工單，等待10分鐘...")
                        time.sleep(600)
                        continue
                        
                    total_orders = len(order_links)
                    log_message(driver, f"找到 {total_orders} 個待處理工單")
                    
                except Exception as e:
                    log_message(driver, f"獲取工單連結失敗: {str(e)}")
                    time.sleep(60)
                    continue
                
                # 處理每個工單
                processed_orders = load_processed_orders()
                for index, order_link in enumerate(order_links, 1):
                    current_order_id = None
                    original_window = driver.current_window_handle
                    try:
                        current_order_id = order_link.text.strip()
                        if not current_order_id:
                            continue
                            
                        if current_order_id in processed_orders:
                            print(f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過,跳過處理")
                            continue
                            
                        log_message(driver, f"開始處理第 {index}/{total_orders} 個工單: {current_order_id}")
                        
                        # 點擊工單連結
                        try:
                            driver.execute_script("arguments[0].click();", order_link)
                            time.sleep(2)
                        except Exception as click_error:
                            log_message(driver, f"點擊工單連結失敗: {str(click_error)}")
                            continue
                            
                        # 處理工單
                        process_order(driver, order_link)
                        
                    except Exception as e:
                        print(f"處理工單 {current_order_id or '未知'} 時發生錯誤: {str(e)}")
                        if driver.current_window_handle != original_window:
                            driver.close()
                            driver.switch_to.window(original_window)
                    finally:
                        time.sleep(2)
                        
                log_message(driver, "所有工單處理完成!")
                time.sleep(300)  # 等待5分鐘後繼續檢查
                    
            except Exception as e:
                log_message(driver, f"處理過程中發生錯誤: {str(e)}")
                time.sleep(60)
                
    except KeyboardInterrupt:
        log_message(driver, "程序被手動中止")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()