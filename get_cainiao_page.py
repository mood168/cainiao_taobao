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
from selenium.common.exceptions import StaleElementReferenceException

def show_notification(driver, message):
    """顯示瀏覽器通知"""
    try:
        # 對消息進行轉義處理，避免 JavaScript 錯誤
        escaped_message = message.replace("'", "\\'").replace("\n", "\\n")
        script = """
        (function() {
            var notification = document.createElement('div');
            notification.textContent = '%s';
            notification.style.cssText = 'position: fixed; top: 20px; right: 20px; background-color: #4CAF50; color: white; padding: 15px; border-radius: 5px; z-index: 9999; max-width: 300px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); opacity: 1; transition: opacity 0.5s;';
            document.body.appendChild(notification);
            setTimeout(function() {
                notification.style.opacity = '0';
                setTimeout(function() {
                    if (notification && notification.parentNode) {
                        notification.parentNode.removeChild(notification);
                    }
                }, 500);
            }, 3000);
        })();
        """ % escaped_message
        driver.execute_script(script)
    except Exception as e:
        print(f"顯示通知時發生錯誤: {str(e)}")
        # 如果顯示通知失敗，至少確保消息被打印出來
        print(f"通知內容: {message}")

def process_current_page(driver, processed_orders, current_page):
    wait = WebDriverWait(driver, 30)
    original_window = driver.current_window_handle
    
    # 重新獲取當前頁面連結(解決頁面切換後元素失效問題)
    current_page_links = retry_operation(
        lambda: driver.find_elements(By.XPATH, "//table//a[string-length(text())=14 and translate(text(), '0123456789', '') = '']"),
        max_retries=5,
        delay=2
    )
    
    if not current_page_links:
        print(f"第 {current_page} 頁沒有找到工單連結")
        return False

    print(f"第 {current_page} 頁找到 {len(current_page_links)} 個工單")
    log_message(driver, f"開始處理第 {current_page} 頁的 {len(current_page_links)} 個工單")

    for index, order_link in enumerate(current_page_links, 1):
        process_single_order(
            driver=driver,
            order_link=order_link,
            processed_orders=processed_orders,
            index=index,
            total_orders=len(current_page_links),
            page_number=current_page
        )
        
        # 每次處理後重新獲取窗口控制權
        driver.switch_to.window(original_window)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        
    return True

def process_single_order(driver, order_link, processed_orders, index, total_orders, page_number):
    wait = WebDriverWait(driver, 30)
    original_window = driver.current_window_handle
    current_order_id = None
    
    try:
        # 驗證元素是否仍有效
        order_link.is_enabled()  
        current_order_id = order_link.text
        order_url = order_link.get_attribute('href')
    except StaleElementReferenceException:
        print(f"第 {page_number} 頁的工單連結已失效，重新獲取...")
        return
    
    try:
        # 檢查是否已處理過
        if current_order_id in processed_orders:
            print(f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過,跳過處理")
            log_message(driver, f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過,跳過處理")
            return
            
        print(f"\n開始處理第 {page_number} 頁的第 {index}/{total_orders} 個工單: {current_order_id}")
        log_message(driver, f"開始處理第 {page_number} 頁的第 {index}/{total_orders} 個工單: {current_order_id}")
        
        # 在開始處理前先保存到記錄中
        save_processed_order(current_order_id, order_url)
        
        # 點擊工單連結並切換到新窗口
        old_handles = driver.window_handles
        driver.execute_script("arguments[0].click();", order_link)
        
        # 等待新窗口出現並切換
        wait.until(lambda d: len(d.window_handles) > len(old_handles))
        new_handle = [h for h in driver.window_handles if h not in old_handles][0]
        driver.switch_to.window(new_handle)
        
        # 等待頁面加載
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        time.sleep(5)
        
        # 確保在"工单基本信息"標籤頁
        info_tab = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@class='next-tabs-tab-inner' and contains(text(), '工单基本信息')]")
        ))
        parent_tab = info_tab.find_element(By.XPATH, "./..")
        if 'active' not in parent_tab.get_attribute('class'):
            info_tab.click()
            time.sleep(2)
        
        # 等待內容區域加載
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'next-tabs-content')))
        time.sleep(2)
        
        # 抓取工單信息
        rows = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, '.next-row')
        ))
        
        # 處理工單信息
        shipmentNo = None
        with open(f'records/{current_order_id}.txt', 'w', encoding='utf-8') as f:
            f.write(f'工单号: {current_order_id}\n')
            for row in rows:
                try:
                    cols = row.find_elements(By.CSS_SELECTOR, '.next-col')
                    for i in range(0, len(cols), 2):
                        if i + 1 < len(cols):
                            field_name = cols[i].text.strip()
                            if '工单基本信息' in field_name:
                                # 處理時間信息
                                time_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', field_name)
                                if time_match:
                                    f.write(f'创建时间: {time_match.group()}\n')
                                # 處理運單號
                                num_parts = field_name.split('运单号')
                                if len(num_parts) > 1:
                                    num_parts_txt = num_parts[1].strip()
                                    shipmentNo = num_parts_txt[:8]
                                    f.write(f'运单号: {shipmentNo}\n')
                                # 處理工單描述
                                description_parts = field_name.split('工单描述')
                                if len(description_parts) > 1:
                                    f.write(f'\n\n工单描述: {description_parts[1]}\n')
                except Exception as e:
                    print(f"處理行數據時發生錯誤: {str(e)}")
                    continue
        
        if shipmentNo:
            # 獲取 API token 並處理貨態
            token = get_api_token()
            if token:
                process_shipment_status(driver, token, shipmentNo, current_order_id)
        
        # 關閉當前窗口並切換回原始窗口
        driver.close()
        driver.switch_to.window(original_window)
        print(f"已完成工單 {current_order_id} 的處理")
        
    except Exception as e:
        print(f"處理工單 {current_order_id} 時發生錯誤: {str(e)}")
        log_message(driver, f"處理工單 {current_order_id} 時發生錯誤: {str(e)}")
        # 確保返回原始窗口
        if driver.current_window_handle != original_window:
            try:
                driver.close()
                driver.switch_to.window(original_window)
            except Exception as close_error:
                print(f"關閉錯誤窗口時發生異常: {str(close_error)}")
    finally:
        time.sleep(2)

def get_api_token():
    token_url = "https://ecapi.sp88.tw/api/Token"
    token_headers = {
        "Content-Type": "application/json",
    }
    token_data = {
        "account": "ESCS",
        "password": "SDG3jdkd59@1"
    }
    try:
        token_response = requests.post(token_url, headers=token_headers, data=json.dumps(token_data), timeout=30)
        if token_response.status_code == 200:
            token_data = token_response.json()
            if token_data.get("token"):
                return token_data.get("token")
    except Exception as e:
        print(f"獲取 API token 時發生錯誤: {str(e)}")
    return None

def process_shipment_status(driver, token, shipmentNo, current_order_id):
    track_url = f"https://ecapi.sp88.tw/api/track/B2C?eshopId=74A&shipmentNo={shipmentNo}"
    track_headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }
    
    try:
        track_response = requests.get(track_url, headers=track_headers)
        if track_response.status_code == 200:
            track_data = track_response.json()
            if track_data.get("status"):
                tracking_info = track_data.get("status")
                process_tracking_result(driver, track_data, current_order_id, tracking_info)
    except Exception as e:
        print(f"處理貨態時發生錯誤: {str(e)}")
        log_message(driver, f"處理貨態時發生錯誤: {str(e)}")

def process_tracking_result(driver, track_data, current_order_id, tracking_info):
    try:
        # 處理日期
        date_str = track_data.get("date")
        try:
            if "T" in date_str:
                memo_date = datetime.strptime(date_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
            else:
                memo_date = datetime.strptime(date_str, "%Y%m%d%H%M%S")
        except:
            memo_date = datetime.now()
        
        # 更新追蹤信息
        tracking_info_new = f'{tracking_info} ({memo_date})'
        errorCode = track_data.get("errorCode")
        
        # 寫入追蹤信息
        with open(f'records/{current_order_id}.txt', 'a', encoding='utf-8') as f:
            f.write(f'\n\n貨態查詢結果: {tracking_info_new}\n')
        
        # 處理結單操作
        if errorCode == 0:
            handle_order_completion(driver, tracking_info, memo_date, current_order_id)
            
    except Exception as e:
        print(f"處理追蹤結果時發生錯誤: {str(e)}")
        log_message(driver, f"處理追蹤結果時發生錯誤: {str(e)}")

def handle_order_completion(driver, tracking_info, memo_date, current_order_id):
    try:
        # 點擊結單按鈕
        finish_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(@class, 'next-btn-helper') and contains(text(), '结单')]")
        ))
        driver.execute_script("arguments[0].click();", finish_btn)
        time.sleep(2)
        
        # 檢查是否有下拉選單
        has_dropdown = False
        try:
            dropdown = WebDriverWait(driver, 2).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span.structFinish-select-trigger')
            ))
            has_dropdown = True
        except:
            pass
        
        if has_dropdown:
            # 處理有下拉選單的情況
            with open(f'需要廠退日_暫不處理單.txt', 'a', encoding='utf-8') as f:
                f.write(f'工單號:{current_order_id} : 需要廠退日_暫不處理單,貨態:{tracking_info}\n')
        else:
            # 處理沒有下拉選單的情況
            handle_no_dropdown_case(driver, tracking_info, memo_date)
            
    except Exception as e:
        print(f"處理結單時發生錯誤: {str(e)}")
        with open(f'投訴單_請人工處理.txt', 'a', encoding='utf-8') as f:
            f.write(f'工單號:{current_order_id} : 投訴單_請人工處理,貨態:{tracking_info}\n')

def handle_no_dropdown_case(driver, tracking_info, memo_date):
    # 關鍵字定義
    preparation_keywords = ["廠商已準備出貨"]
    oversize_keywords = ["包裹材積超過規範", "將退回廠商"]
    store_issue_keywords = ["因門市因素無法配送，請與賣方客服聯繫重選取件門市", "請與賣方客服聯繫"]
    
    # 根據不同情況設置回覆訊息
    if any(keyword in tracking_info for keyword in preparation_keywords):
        tracking_info_new = "我方未收到包裹，請與菜鳥台灣倉確認，感謝"
    elif any(keyword in tracking_info for keyword in oversize_keywords):
        return_future_date = (memo_date + timedelta(days=5)).strftime("%m/%d")
        tracking_info_new = f"包裹材積超過規範，預計{return_future_date}退回清關行, 謝謝"
    elif any(keyword in tracking_info for keyword in store_issue_keywords):
        issue_future_date = (memo_date + timedelta(days=7)).strftime("%m/%d")
        tracking_info_new = f"門市關轉，預計{issue_future_date}退回清關行, 謝謝"
    else:
        tracking_info_new = tracking_info
    
    # 填寫回覆
    memo_textarea = WebDriverWait(driver, 2).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'textarea[name="memo"]')
    ))
    actions = webdriver.ActionChains(driver)
    actions.click(memo_textarea).perform()
    pyperclip.copy(tracking_info_new)
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    time.sleep(1)
    
    # 提交
    submit_btn = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '确定并提交')]"))
    )
    driver.execute_script("arguments[0].click();", submit_btn)
    time.sleep(2)
    
    # 等待對話框消失
    try:
        WebDriverWait(driver, 10).until_not(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'next-dialog-wrapper')]"))
        )
        time.sleep(2)
    except Exception as e:
        print(f"等待對話框消失時發生錯誤: {str(e)}")

def log_message(driver, message):
    """同時執行 print 和瀏覽器通知"""
    print(message)
    if driver:
        show_notification(driver, message)

# 設置Chrome選項
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')  # 忽略SSL證書錯誤
chrome_options.add_argument('--ignore-ssl-errors')  # 忽略SSL錯誤
chrome_options.add_argument('--window-size=1920,1080')  # 設置窗口大小
chrome_options.add_argument('--disable-infobars')  # 禁用infobars
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_experimental_option("prefs", {
    "profile.password_manager_enabled": False,
    "credentials_enable_service": False,
})
# chrome_options.add_argument('--headless')  # 無頭模式,調試時先註釋掉

driver = None
try:
    print("開始執行自動化流程...")
    
    # 初始化WebDriver時添加選項
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),  # 自動管理驅動版本
        options=chrome_options
    )
    wait = WebDriverWait(driver, 30)  # 增加默認等待時間
    original_window = driver.current_window_handle
    print("瀏覽器已啟動")
    log_message(driver, "瀏覽器已啟動")
    
    driver.get('https://desk.cainiao.com/')
    print("已訪問初始頁面")
    log_message(driver, "已訪問初始頁面")
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

    # 等待並切換到iframe
    def switch_to_iframe():
        iframe = wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, 'alibaba-login-box'))
        )
        print("已切換到登入iframe")
        return iframe

    retry_operation(switch_to_iframe)

    # 輸入登入信息
    def input_credentials():
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

    retry_operation(input_credentials)

    # 切換回主文檔
    driver.switch_to.default_content()
    print("已切換回主文檔")
    
    # 等待登入完成
    print("等待登入完成...")
    log_message(driver, "等待登入完成...")
    def wait_for_login_success():
        def check_login():
            current_url = driver.current_url
            print(f"當前URL: {current_url}")
            if "login" not in current_url and "desk.cainiao.com" in current_url:
                print("登入成功!")
                log_message(driver, "登入成功!")
                return True
            return False

        max_retries = 2
        for i in range(max_retries):
            try:
                if check_login():
                    return True
                time.sleep(3)
                print(f"等待登入完成... ({i+1}/{max_retries})")
            except Exception as e:
                print(f"檢查登入狀態時發生錯誤: {e}")
            
        return False

    if not wait_for_login_success():
        log_message(driver, "登入超時")
        raise Exception("登入超時")

    # 導向到任務頁面
    target_url = 'https://desk.cainiao.com/unified/myTask/pendingTask'
    print(f"正在導向到任務頁面: {target_url}")
    def navigate_to_task_page():
        driver.get(target_url)
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        # 等待表格元素出現
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        time.sleep(5)  # 額外等待動態內容

    retry_operation(navigate_to_task_page, max_retries=3, delay=5)
    
        
    # 確保列表已完全載入
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
    time.sleep(3)  # 額外等待確保列表穩定
    
    print("開始尋找工單號連結...")
    log_message(driver, "開始尋找工單號連結...")
    def find_order_links():
        # 等待表格完全加載
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        
        all_order_links = []  # 存儲所有頁面的工單連結
        
        while True:
            # 使用多種定位策略獲取當前頁面的工單連結
            current_page_links = None
            try:
                # 方法1: 使用XPath精確匹配14位數字
                current_page_links = driver.find_elements(By.XPATH, "//table//a[string-length(text())=14 and translate(text(), '0123456789', '') = '']")
            except:
                try:
                    # 方法2: 使用CSS選擇器
                    current_page_links = driver.find_elements(By.CSS_SELECTOR, "table tbody tr td:first-child a")
                except:
                    # 方法3: 使用更寬鬆的XPath
                    current_page_links = driver.find_elements(By.XPATH, "//*[contains(@class, 'ticket-number') or contains(@class, 'order-number')]")
            
            if not current_page_links:
                if not all_order_links:  # 如果是第一頁就沒有連結,則報錯
                    log_message(driver, "找不到工單號連結")
                    raise Exception("找不到工單號連結")
                break  # 如果不是第一頁,則表示已經處理完所有頁面
            
            # 將當前頁面的連結添加到總列表中
            all_order_links.extend(current_page_links)
            
            # 尋找下一頁按鈕
            try:
                # 找到當前頁碼
                current_page = driver.find_element(By.XPATH, "//button[contains(@class, 'next-pagination-item next-current')]//span").text
                next_page = str(int(current_page) + 1)
                
                # 尋找下一頁按鈕
                next_page_btn = driver.find_element(By.XPATH, f"//button[contains(@class, 'next-pagination-item')]//span[text()='{next_page}']")
                
                # 點擊下一頁
                driver.execute_script("arguments[0].click();", next_page_btn)
                print(f"正在切換到第 {next_page} 頁...")
                log_message(driver, f"正在切換到第 {next_page} 頁...")
                
                # 等待新頁面加載
                time.sleep(3)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            except:
                print("已經是最後一頁或切換頁面失敗")
                break
        
        # 移除重複的工單連結
        unique_order_links = []
        seen_order_ids = set()
        
        for link in all_order_links:
            order_id = link.text.strip()
            if order_id not in seen_order_ids:
                seen_order_ids.add(order_id)
                unique_order_links.append(link)
        
        order_links = unique_order_links
        print(f"總共找到 {len(order_links)} 個工單號連結")
        log_message(driver, f"總共找到 {len(order_links)} 個工單號連結")
        return order_links

    def load_processed_orders():
        try:
            if os.path.exists('processed_orders.json'):
                with open('processed_orders.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # print("\n已處理的工單列表:")
                    # for order_id, info in data.items():
                    #     print(f"工單號: {order_id}, 處理時間: {info['processed_time']}")
                    return data
            else:
                print("\n未找到處理記錄文件,將創建新的記錄")
                log_message(driver, "未找到處理記錄文件,將創建新的記錄")
                return {}
        except Exception as e:
            print(f"讀取處理記錄時發生錯誤: {str(e)}")
            print("將重新創建處理記錄")
            log_message(driver, "將重新創建處理記錄")
            return {}

    def save_processed_order(order_id, order_url):
        try:
            processed_orders = load_processed_orders()
            if order_id not in processed_orders:  # 只有在不存在時才添加
                processed_orders[order_id] = {
                    'url': order_url,
                    'processed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                with open('processed_orders.json', 'w', encoding='utf-8') as f:
                    json.dump(processed_orders, f, ensure_ascii=False, indent=2)
                print(f"已將工單 {order_id} 添加到處理記錄")
                log_message(driver, f"已將工單 {order_id} 添加到處理記錄")
            else:
                print(f"工單 {order_id} 已存在於處理記錄中")
                log_message(driver, f"工單 {order_id} 已存在於處理記錄中")
        except Exception as e:
            print(f"保存處理記錄時發生錯誤: {str(e)}")
            print("錯誤詳情:", traceback.format_exc())
            log_message(driver, f"保存處理記錄時發生錯誤: {str(e)}\n錯誤詳情: {traceback.format_exc()}")

    # 獲取所有工單連結
    order_links = retry_operation(find_order_links, max_retries=10, delay=3)
    total_orders = len(order_links)
    print(f"總共有 {total_orders} 個工單需要處理")


    # 載入已處理的工單記錄
    processed_orders = load_processed_orders()
    print(f"已載入處理記錄,共 {len(processed_orders)} 個工單")
    # 檢查並建立records資料夾
    if not os.path.exists('records'):
        os.makedirs('records')
        print("已建立records資料夾")
        log_message(driver, "已建立records資料夾")
    # 遍歷處理每個工單
    for index, order_link in enumerate(order_links, 1):
        current_order_id = None
        try:
            current_order_id = order_link.text
            order_url = order_link.get_attribute('href')
            
            # 檢查是否已處理過
            if current_order_id in processed_orders:
                print(f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過,跳過處理,繼續下一個工單")
                log_message(driver, f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過,跳過處理,繼續下一個工單")
                continue
                
            print(f"\n開始處理第 {index}/{total_orders} 個工單: {current_order_id}")
            log_message(driver, f"\n開始處理第 {index}/{total_orders} 個工單: {current_order_id}")
            print(f"工單URL: {order_url}")
            
            # 在開始處理前先保存到記錄中,避免重複處理
            save_processed_order(current_order_id, order_url)
            
            # 點擊工單連結並切換到新窗口
            original_window = driver.current_window_handle
            old_handles = driver.window_handles
            
            # 使用JavaScript點擊
            driver.execute_script("arguments[0].click();", order_link)
            print("已點擊工單連結")
            log_message(driver, "已點擊工單連結")
            try:
                # 等待新窗口出現
                wait.until(lambda d: len(d.window_handles) > len(old_handles))
                new_handle = [h for h in driver.window_handles if h not in old_handles][0]
                print(f"切換到新窗口: {new_handle}")
                
                # 切換到新窗口
                driver.switch_to.window(new_handle)
                
                try:
                    # 等待頁面加載完成
                    wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
                    time.sleep(5)  # 增加等待時間
                    
                    # 確保在"工单基本信息"標籤頁
                    info_tab = wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//div[@class='next-tabs-tab-inner' and contains(text(), '工单基本信息')]")
                    ))
                    print("找到工单基本信息標籤")
                    # 檢查是否需要點擊標籤
                    parent_tab = info_tab.find_element(By.XPATH, "./..")
                    if 'active' not in parent_tab.get_attribute('class'):
                        info_tab.click()
                        print("點擊工单基本信息標籤")
                        log_message(driver, "點擊工单基本信息標籤")
                        time.sleep(2)
                    
                    print("已切換到工单基本信息標籤頁")
                    
                    # 等待內容區域加載
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'next-tabs-content')))
                    time.sleep(2)  # 額外等待內容加載
                    
                    # 抓取工單信息
                    field_data = {}
                    
                    # 等待並獲取所有字段行
                    rows = wait.until(EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, '.next-row')
                    ))
                    print(f"找到 {len(rows)} 行數據")
                    log_message(driver, f"找到 {len(rows)} 行數據")
                    # 創建文件並寫入工單號
                    with open(f'records/{current_order_id}.txt', 'w', encoding='utf-8') as f:
                        f.write(f'工单号: {current_order_id}\n')
                        # 遍歷每一行尋找指定字段
                        for row in rows:
                            try:
                                cols = row.find_elements(By.CSS_SELECTOR, '.next-col')
                                for i in range(0, len(cols), 2):
                                    if i + 1 < len(cols):
                                        field_name = cols[i].text.strip()
                                        field_value = cols[i + 1].text.strip()
                                        if '工单基本信息' in field_name:
                                            time_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', field_name)
                                            # 使用正則表達式匹配時間格式 YYYY-MM-DD hh:mm:ss
                                            if time_match:
                                                f.write(f'创建时间: {time_match.group()}\n')
                                            num_parts = field_name.split('运单号')
                                            if len(num_parts) > 1:
                                                num_parts_txt = num_parts[1].strip()
                                                f.write(f'运单号: {num_parts_txt[:8]}\n')
                                                log_message(driver, f"運單號: {num_parts_txt[:8]}")
                                                shipmentNo = num_parts_txt[:8]
                                            description_parts = field_name.split('工单描述')
                                            if len(description_parts) > 1:
                                                f.write(f'\n\n工单描述: {description_parts[1]}\n')
                                        if '当前工单记录' in field_name:
                                            record_parts = field_name.split('当前工单记录')
                                            if len(record_parts) > 1:
                                                # 計算當前工單記錄中的時間戳數量
                                                all_matches = re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', record_parts[1])
                                                match_count = len(list(all_matches))
                                                # 找到所有時間戳並添加換行符
                                                for post_time_match in re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', record_parts[1]):
                                                    record_parts[1] = record_parts[1].replace(post_time_match.group(), post_time_match.group() + "\n")
                                                    
                                                f.write(f'\n\n当前工单记录: {record_parts[1]}\n')

                            except Exception as e:
                                print(f"處理行數據時發生錯誤: {str(e)}")
                                continue
                    # 獲取 API token
                    try:
                        token_url = "https://ecapi.sp88.tw/api/Token"
                        token_headers = {
                            "Content-Type": "application/json",
                        }
                        token_data = {
                            "account": "ESCS",
                            "password": "SDG3jdkd59@1"
                        }
                        token_data_str = json.dumps(token_data)
                        token = None  # 初始化 token 變量
                        token_success = False  # 添加標誌變量
                        
                        # print(f"正在請求 token, URL: {token_url}")
                        # print(f"請求數據: {token_data_str}")
                        
                        try:
                            token_response = requests.post(token_url, headers=token_headers, data=token_data_str, timeout=30)
                            # print(f"API 響應狀態碼: {token_response.status_code}")
                            # print(f"API 響應內容: {token_response.text}")
                            
                            if token_response.status_code == 200:
                                token_data = token_response.json()
                                if token_data.get("token") != None:
                                    token = token_data.get("token")
                                    token_success = True  # 設置成功標誌
                                    # print("成功獲取 token")
                                else:
                                    error_msg = token_data.get("Message", "未知錯誤")
                                    # print(f"獲取 token 失敗,API 返回錯誤: {error_msg}")
                                    # print(f"完整響應: {token_data}")
                            else:
                                print(f"HTTP 請求失敗,狀態碼: {token_response.status_code}")
                                print(f"錯誤響應: {token_response.text}")
                        except requests.exceptions.Timeout:
                            print("請求超時,請檢查網絡連接或 API 服務器狀態")
                        except requests.exceptions.RequestException as e:
                            print(f"請求異常: {str(e)}")
                        except json.JSONDecodeError as e:
                            print(f"解析 JSON 響應失敗: {str(e)}")
                            print(f"原始響應內容: {token_response.text}")
                        
                        # 只有在成功獲取 token 時才繼續執行
                        if token_success:
                            # 使用獲取的 token 調用追蹤 API
                            track_url = f"https://ecapi.sp88.tw/api/track/B2C?eshopId=74A&shipmentNo={shipmentNo}"  # 修正字符串插值語法
                            track_headers = {
                                "Content-Type": "application/json", 
                                "Authorization": f"Bearer {token}"
                            }
                            
                            print(f"正在請求追蹤信息, URL: {track_url}")
                            track_response = None
                            track_response = requests.get(track_url, headers=track_headers)
                            print(f"追蹤 API 響應狀態碼: {track_response.status_code}")
                            print(f"追蹤 API 響應內容: {track_response.text}")
                            
                            if track_response.status_code == 200:
                                track_data = track_response.json()
                                if track_data.get("status") != None:
                                    tracking_info = track_data.get("status")
                                    log_message(driver, f'貨態: {tracking_info}')
                                    # 修改日期格式處理
                                    date_str = track_data.get("date")
                                    try:
                                        if "T" in date_str:
                                            # 處理包含 'T' 的ISO格式
                                            memo_date = datetime.strptime(date_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
                                        else:
                                            # 處理標準格式
                                            memo_date = datetime.strptime(date_str, "%Y%m%d%H%M%S")
                                    except Exception as e:
                                        print(f"日期格式轉換錯誤: {str(e)}")
                                        memo_date = datetime.now()
                                    
                                    errorCodeDescription = track_data.get("errorCodeDescription")
                                    errorCode = track_data.get("errorCode")
                                    # 將追蹤信息寫入文件
                                    # log_message(driver, f'(自動寫入)新留言')
                                    with open(f'records/{current_order_id}.txt', 'a', encoding='utf-8') as f:
                                        f.write(f'(自動寫入)新留言: {tracking_info}\n{memo_date}')
                                        f.write(f'\n\n貨態查詢:')
                                        if errorCode == 0:
                                            f.write(f'\n結果: {track_response.text.replace('"', '').replace('\n', '')}\n')
                                        else:
                                            f.write(f'\n結果: {errorCodeDescription.replace('"', '')}\n')
                                    
                                    # 點擊結單按鈕
                                    # 只有在 errorCode = 0 時才執行結單操作
                                    if errorCode == 0:
                                        try:
                                            print("準備點擊結單按鈕...")
                                            # 使用更短的等待時間尋找結單按鈕
                                            
                                            finish_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                                                (By.XPATH, "//span[contains(@class, 'next-btn-helper') and contains(text(), '结单')]")
                                            ))
                                            driver.execute_script("arguments[0].click();", finish_btn)
                                            print("已點擊結單按鈕")

                                            # 等待結單對話框出現
                                            time.sleep(2)  # 等待對話框完全顯示
                                            # 檢查是否有下拉選單
                                            has_dropdown = False
                                            try:
                                                dropdown = WebDriverWait(driver, 2).until(EC.presence_of_element_located(
                                                    (By.CSS_SELECTOR, 'span.structFinish-select-trigger')
                                                ))
                                                has_dropdown = True
                                                print("檢測到有下拉選單")   
                                                log_message(driver, "檢測到有下拉選單")
                                            except:
                                                print("檢測到沒有下拉選單")
                                                log_message(driver, "檢測到沒有下拉選單")
                                            tracking_info_new = f'{tracking_info} ({memo_date})'
                                            # # 根據tracking_info選擇對應選項
                                            delivery_keywords = ["包裹配達門市", "包裹已成功取件", "包裹已送達物流中心，進行理貨中", "等待配送中", "包裹配送中"]
                                            preparation_keywords = ["廠商已準備出貨"]
                                            force_majeure_keywords = ["持續未配送", "取件門市位於離島地區"]
                                            return_keywords = ["包裹逾期未領，已送回物流中心", "包裹已送達物流中心，進行退貨處理中"]
                                            oversize_keywords = ["包裹材積超過規範", "將退回廠商"]
                                            store_issue_keywords = ["因門市因素無法配送，請與賣方客服聯繫重選取件門市", "請與賣方客服聯繫"]
                                            # 根據不同情況處理
                                            if has_dropdown:
                                                with open(f'需要廠退日_暫不處理單.txt', 'a', encoding='utf-8') as f:
                                                    f.write(f'工單號:{current_order_id} : 需要廠退日_暫不處理單,貨態:{tracking_info_new}\n')
                                                # # 點擊下拉選單
                                                # dropdown.click()
                                                # time.sleep(1)
                                                # log_message(driver, "已點擊下拉選單")
                                                
                                                
                                                # # 設置默認回覆訊息
                                                # reply_message = tracking_info
                                                # option_text = ""
                                                
                                                # # 根據關鍵字選擇選項和設置回覆訊息
                                                # if any(keyword in tracking_info for keyword in preparation_keywords):
                                                #     option_text = "包裹实际未交接、未收到包裹"
                                                #     reply_message = "我方未收到包裹，請與菜鳥台灣倉確認，感謝"
                                                # elif any(keyword in tracking_info for keyword in force_majeure_keywords):
                                                #     option_text = "不可抗力已报备"
                                                #     reply_message = "因取件門市位於離島地區，船班需視當地海象氣候配送，包裹到店將發送簡訊通知，還請以到店簡訊通知為主，造成不便，敬請見諒，感謝"
                                                # elif any(keyword in tracking_info for keyword in return_keywords):
                                                #     option_text = "无法物流履约，不需要菜鸟协助"
                                                #     future_date = (datetime.now() + timedelta(days=7)).strftime("%m/%d")
                                                #     reply_message = f"將退回清關行、具体原因：逾期未取、天猫海外预计时间：{future_date}"
                                                # elif any(keyword in tracking_info for keyword in store_issue_keywords):
                                                #     option_text = "无法物流履约，不需要菜鸟协助"
                                                #     future_date = (datetime.now() + timedelta(days=7)).strftime("%m/%d")
                                                #     reply_message = f"將退回清關行、具体原因：門市關轉、天猫海外预计时间：{future_date}"
                                                # else:
                                                #     #寫入記錄
                                                #     with open(f'records/需要廠退日_暫不處理單.txt', 'a', encoding='utf-8') as f:
                                                #         f.write(f'工單號:{current_order_id} : 需要廠退日_暫不處理單,貨態:{tracking_info},貨態時間:{memo_date}\n')
                                                
                                                # if option_text:
                                                #     # 選擇下拉選單選項
                                                #     option = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((
                                                #         By.XPATH, f"//span[contains(text(), '{option_text}')]"
                                                #     )))
                                                #     driver.execute_script("arguments[0].click();", option)
                                                #     print(f"已選擇選項: {option_text}")
                                                
                                                # # 找到備註文本框
                                                # memo_textarea = WebDriverWait(driver, 2).until(EC.presence_of_element_located(
                                                #     (By.CSS_SELECTOR, 'textarea[name="memo"]')
                                                # ))
                                                
                                                # # 使用剪貼板貼上文字
                                                # pyperclip.copy(reply_message)  # 將文字複製到剪貼板
                                                # actions = webdriver.ActionChains(driver)
                                                # actions.click(memo_textarea).perform()  # 點擊文本框
                                                # memo_textarea.clear()  # 清空文本框
                                                # actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()  # 貼上
                                                # time.sleep(0.5)
                                                
                                                # # 點擊確定按鈕
                                                # confirm_btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((
                                                #     By.XPATH, "//span[contains(@class, 'structFinish-btn-helper') and contains(text(), '确定')]"
                                                # )))
                                            else:
                                                # 沒有下拉選單的情況
                                                memo_textarea = WebDriverWait(driver, 2).until(EC.presence_of_element_located(
                                                    (By.CSS_SELECTOR, 'textarea[name="memo"]')
                                                ))
                                                # preparation_keywords = ["廠商已準備出貨"]
                                                if any(keyword in tracking_info for keyword in preparation_keywords):
                                                    tracking_info_new = "我方未收到包裹，請與菜鳥台灣倉確認，感謝"
                                                return_future_date = (memo_date + timedelta(days=5)).strftime("%m/%d")
                                                if any(keyword in tracking_info for keyword in oversize_keywords):
                                                        tracking_info_new = f"包裹材積超過規範，預計{return_future_date}退回清關行, 謝謝"
                                                issue_future_date = (memo_date + timedelta(days=7)).strftime("%m/%d")
                                                if any(keyword in tracking_info for keyword in store_issue_keywords):
                                                    tracking_info_new = f"門市關轉，預計{issue_future_date}退回清關行, 謝謝"
                                                
                                                
                                                # 檢查是否為結單對話框或留言對話框
                                            
                                                # 使用剪貼板貼上文字
                                                pyperclip.copy(tracking_info_new)  # 將文字複製到剪貼板
                                                actions = webdriver.ActionChains(driver)
                                                actions.click(memo_textarea).perform()  # 點擊文本框
                                                # 清空文本框
                                                actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()  # 貼上
                                                time.sleep(1)
                                                finish_btn = driver.find_element(By.XPATH, "//span[contains(@class, 'cDeskStructFunctionComponent-btn-helper') and contains(text(), '确定并提交')]")
                                                submit_btn_text = "确定并提交"                                               
                                                
                                                
                                                submit_btn = WebDriverWait(driver, 2).until(
                                                    EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(), '{submit_btn_text}')]"))
                                                )
                                                driver.execute_script("arguments[0].click();", submit_btn)
                                                time.sleep(2)
                                                print(f"已點擊{submit_btn_text}按鈕")
                                                
                                                # 等待結單對話框消失
                                                try:
                                                    WebDriverWait(driver, 10).until_not(
                                                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'next-dialog-wrapper')]"))
                                                    )
                                                    time.sleep(2)
                                                    print("對話框已消失")
                                                    # log_message(driver, "對話框已消失")
                                                except Exception as e:
                                                    print(f"等待對話框消失時發生錯誤: {str(e)}")
                                                    log_message(driver, f"等待對話框消失時發生錯誤: {str(e)}")
                                                time.sleep(2)
                                                
                                        except Exception as e:
                                            with open(f'投訴單_請人工處理.txt', 'a', encoding='utf-8') as f:
                                                f.write(f'工單號:{current_order_id} : 投訴單_請人工處理,貨態:{tracking_info_new}\n')
                                            print(f"此為投訴單或此頁面無結單按鈕: {str(e)}")
                                            log_message(driver, f"此為投訴單或此頁面無結單按鈕: {str(e)},關閉視窗並繼續下一筆...")
                                            print("關閉視窗並繼續下一筆...")
                                            time.sleep(2)
                                            # 關閉當前視窗並切換回原始視窗
                                            try:
                                                driver.close()
                                                driver.switch_to.window(original_window)
                                                print("已關閉視窗並切換回原始視窗")
                                                log_message(driver, "已關閉視窗並切換回原始視窗")
                                            except Exception as close_error:
                                                print(f"關閉視窗時發生錯誤: {str(close_error)}")
                                                log_message(driver, f"關閉視窗時發生錯誤: {str(close_error)}")
                                            continue
                                    else:
                                        print(f"貨態查詢失敗 (errorCode: {errorCode}), 跳過結單操作")
                                        log_message(driver, f"貨態查詢失敗 (errorCode: {errorCode}), 跳過結單操作")
                                        with open(f'貨態查詢失敗_無法處理單.txt', 'a', encoding='utf-8') as f:
                                            f.write(f'工單號:{current_order_id} : 錯誤描述: {errorCodeDescription}\n')
                                        # 關閉當前視窗並切換回原始視窗
                                        try:
                                            driver.close()
                                            driver.switch_to.window(original_window)
                                            print("已關閉視窗並切換回原始視窗")
                                            log_message(driver, "已關閉視窗並切換回原始視窗")
                                        except Exception as close_error:
                                            print(f"關閉視窗時發生錯誤: {str(close_error)}")
                                            log_message(driver, f"關閉視窗時發生錯誤: {str(close_error)}")
                                        continue
                    except Exception as e:
                        print(f"調用 API 時發生錯誤: {str(e)}")
                        log_message(driver, f"調用 API 時發生錯誤: {str(e)}")
                    print(f"已完成數據寫入到 {current_order_id}.txt")
                    log_message(driver, f"已完成數據寫入到 {current_order_id}.txt")
                except Exception as e:
                    print(f"處理工單內容時發生錯誤: {str(e)}")
                    log_message(driver, f"處理工單內容時發生錯誤: {str(e)}")
                    raise  # 重新拋出異常以觸發外層的清理代碼
                    
            except Exception as e:
                print(f"處理新窗口時發生錯誤: {str(e)}")
                log_message(driver, f"處理新窗口時發生錯誤: {str(e)}")
                raise  # 重新拋出異常以觸發外層的清理代碼
                
            # 在成功處理後保存記錄
            try:
                save_processed_order(current_order_id, order_url)
                print(f"已記錄工單 {current_order_id} 的處理狀態")
                log_message(driver, f"已記錄工單 {current_order_id} 的處理狀態")
            except Exception as e:
                print(f"保存工單處理記錄時發生錯誤: {str(e)}")
                log_message(driver, f"保存工單處理記錄時發生錯誤: {str(e)}")
                
        except Exception as e:
            print(f"處理工單 {current_order_id or '未知'} 時發生錯誤: {str(e)}")
            log_message(driver, f"處理工單 {current_order_id or '未知'} 時發生錯誤: {str(e)}")
            if driver.current_window_handle != original_window:
                try:
                    driver.close()
                    driver.switch_to.window(original_window)
                except Exception as close_error:
                    print(f"關閉錯誤窗口時發生異常: {str(close_error)}")
                    log_message(driver, f"關閉錯誤窗口時發生異常: {str(close_error)}")
        else:
            # 正常完成時的清理代碼
            try:
                driver.close()
                driver.switch_to.window(original_window)
                print(f"已完成工單 {current_order_id} 的處理並關閉窗口")
            except Exception as close_error:
                print(f"關閉窗口時發生錯誤: {str(close_error)}")
        finally:
            time.sleep(2)  # 無論成功與否都等待一下再處理下一個
            
    print("\n所有工單處理完成!")
    log_message(driver, "\n所有工單處理完成!")

    def reload_and_check():
        try:
            print("\n重新載入頁面...")
            driver.refresh()
            # 等待頁面加載完成
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
            # 等待表格元素出現
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            time.sleep(5)  # 額外等待動態內容
            
            # 重新獲取工單連結
            order_links = retry_operation(find_order_links, max_retries=10, delay=3)
            total_orders = len(order_links)
            if total_orders > 0:
                print(f"檢查到 {total_orders} 個新工單")
            return total_orders, order_links
            
        except Exception as e:
            print(f"重新載入頁面時發生錯誤: {str(e)}")
            return 0, []

    # 主循環：持續檢查和處理工單
    while True:
        try:
            print("開始檢查工單...")
            log_message(driver, "開始檢查工單...")
            
            # 重新載入頁面
            driver.get(target_url)
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            time.sleep(5)  # 等待動態內容載入
            
            has_next_page = True
            current_page = 1
            
            while has_next_page:
                print(f"\n正在處理第 {current_page} 頁的工單...")
                log_message(driver, f"\n正在處理第 {current_page} 頁的工單...")
                
                # 強制重新獲取頁面元素
                driver.execute_script("location.reload(true);")  # 強制硬刷新
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
                time.sleep(5)
                
                # 使用新版處理函式
                page_has_content = process_current_page(driver, processed_orders, current_page)
                
                if not page_has_content:
                    has_next_page = False
                    break
                
                # 切換頁面前確認所有窗口已關閉
                if len(driver.window_handles) > 1:
                    for handle in driver.window_handles[1:]:
                        driver.switch_to.window(handle)
                        driver.close()
                    driver.switch_to.window(original_window)
                
                # 頁面切換邏輯加強(約1040-1054行)
                try:
                    next_page = current_page + 1
                    next_page_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, f"//button[contains(@class, 'next-pagination-item')]//span[text()='{next_page}']"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView(true);", next_page_btn)
                    driver.execute_script("arguments[0].click();", next_page_btn)
                    
                    # 新增頁面切換驗證
                    WebDriverWait(driver, 15).until(
                        lambda d: d.find_element(By.XPATH, f"//li[@title='{next_page}']").get_attribute("class") == "next-pagination-item next-active"
                    )
                    current_page = next_page
                    time.sleep(5)
                except Exception as e:
                    print(f"頁面切換失敗: {str(e)}，可能已是最後一頁")
                    has_next_page = False
            
            print("\n所有頁面的工單處理完成,等待5分鐘後重新檢查...")
            log_message(driver, "\n所有頁面的工單處理完成,等待5分鐘後重新檢查...")
            time.sleep(300)
            
        except Exception as e:
            print(f"檢查工單時發生錯誤: {str(e)}")
            log_message(driver, f"檢查工單時發生錯誤: {str(e)}")
            print("等待5分鐘後重試...")
            log_message(driver, "等待5分鐘後重試...")
            time.sleep(300)
            continue
    
except KeyboardInterrupt:
    print("\n檢測到Ctrl+C，正在退出程式...")
except Exception as e:
    print(f"發生錯誤: {e}")
    print("錯誤詳情:")
    traceback.print_exc()
finally:
    if driver:
        driver.quit()
        print("瀏覽器已關閉")