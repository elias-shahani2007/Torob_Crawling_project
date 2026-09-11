from datetime import datetime
import json
import os
import random
import sys
import threading
import time
from tkinter import filedialog, messagebox
import urllib.parse
import customtkinter as ctk
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
from playwright.sync_api import sync_playwright

try:
  from torob_chart import TorobChartModule
except ImportError:
  TorobChartModule = None


def get_resource_path(relative_path):
  if hasattr(sys, '_MEIPASS'):
    return os.path.join(sys._MEIPASS, relative_path)
  return os.path.join(os.path.abspath('.'), relative_path)


ctk.set_appearance_mode('Dark')
ctk.set_default_color_theme('blue')


class TorobScraperApp(ctk.CTk):
      max_results = 9
      def __init__(self):
        super().__init__()

        self.title('سامانه استخراج داده ترب')
        self.geometry('720x760')
        self.resizable(False, False)

        try:
          icon_path = get_resource_path('icon.ico')
          if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        except Exception:
          pass

        self.is_running = False
        self.stop_requested = False
        self.selected_file_path = ''

        self.create_widgets()

      def create_widgets(self):
        self.lbl_title = ctk.CTkLabel(
            self,
            text='📊 استخراج‌گر هوشمند اطلاعات ترب',
            font=ctk.CTkFont(family='Vazirmatn', size=22, weight='bold'),
        )
        self.lbl_title.pack(pady=(15, 10))

        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(padx=20, pady=8, fill='x')

        self.btn_select_file = ctk.CTkButton(
            self.frame_file,
            text='انتخاب فایل اکسل ورودی',
            command=self.select_file,
            width=180,
        )
        self.btn_select_file.pack(side='right', padx=10, pady=10)

        self.lbl_file_path = ctk.CTkLabel(
            self.frame_file, text='هیچ فایلی انتخاب نشده است', text_color='gray'
        )
        self.lbl_file_path.pack(side='left', padx=10, pady=10)

        self.frame_rows = ctk.CTkFrame(self)
        self.frame_rows.pack(padx=20, pady=8, fill='x')

        self.lbl_end = ctk.CTkLabel(self.frame_rows, text='تا سطر:')
        self.lbl_end.pack(side='right', padx=(2, 8), pady=10)

        self.entry_end_row = ctk.CTkEntry(self.frame_rows, width=60)
        self.entry_end_row.insert(0, '10')
        self.entry_end_row.pack(side='right', padx=(0, 10), pady=10)

        self.lbl_start = ctk.CTkLabel(self.frame_rows, text='از سطر:')
        self.lbl_start.pack(side='right', padx=(2, 5), pady=10)

        self.entry_start_row = ctk.CTkEntry(self.frame_rows, width=60)
        self.entry_start_row.insert(0, '1')
        self.entry_start_row.pack(side='right', padx=(0, 10), pady=10)

        self.switch_headless = ctk.CTkSwitch(
            self.frame_rows, text='مرورگر مخفی'
        )
        self.switch_headless.pack(side='left', padx=10, pady=10)
        self.switch_headless.deselect()

        self.frame_actions = ctk.CTkFrame(self)
        self.frame_actions.pack(padx=20, pady=8, fill='x')

        self.btn_start = ctk.CTkButton(
            self.frame_actions,
            text='▶ شروع استخراج',
            fg_color='#28a745',
            hover_color='#218838',
            command=self.start_scraping_thread,
            height=40,
            font=ctk.CTkFont(size=14, weight='bold'),
        )
        self.btn_start.pack(side='right', padx=10, pady=10, expand=True, fill='x')

        self.btn_stop = ctk.CTkButton(
            self.frame_actions,
            text='🚨 توقف و ذخیره اضطراری',
            fg_color='#dc3545',
            hover_color='#c82333',
            command=self.trigger_emergency_stop,
            height=40,
            state='disabled',
            font=ctk.CTkFont(size=14, weight='bold'),
        )
        self.btn_stop.pack(side='left', padx=10, pady=10, expand=True, fill='x')

        self.txt_log = ctk.CTkTextbox(self, width=680, height=350)
        self.txt_log.pack(padx=20, pady=(8, 15))
        self.log('💡 آماده‌باش: فایل اکسل را انتخاب کرده و «شروع استخراج» را بزنید.')

      def log(self, message):
        self.txt_log.insert('end', f'{message}\n')
        self.txt_log.see('end')

      def select_file(self):
        file_path = filedialog.askopenfilename(
            title='انتخاب فایل اکسل محصولات',
            filetypes=[('Excel Files', '*.xlsx *.xls')],
        )
        if file_path:
          self.selected_file_path = file_path
          filename = os.path.basename(file_path)
          self.lbl_file_path.configure(text=filename, text_color='white')
          self.log(f'📁 فایل ورودی انتخاب شد: {filename}')

      def trigger_emergency_stop(self):
        if self.is_running:
          self.stop_requested = True
          self.log('\n🚨 درخواست توقف دریافت شد... ذخیره‌سازی انجام می‌شود.')

      def handle_captcha_if_present(self, page):
        """بررسی هوشمند کپچا فقط از روی عنوان تب مرورگر و URL اصلی (برطرف شدن توهم کپچا)"""
        try:
          title = page.title().lower()
          current_url = page.url.lower()

          is_blocked = (
              'captcha' in current_url
              or 'challenge' in current_url
              or '403' in title
              or 'forbidden' in title
              or 'تایید هویت' in title
              or 'دسترسی غیرمجاز' in title
          )

          if is_blocked:
            self.log('\n⚠️ [سد امنیتی ترب] کپچای واقعی ظاهر شد!')
            self.log('👉 لطفاً کپچا را در مرورگر حل کنید...')

            messagebox.showwarning(
                'تایید هویت',
                'لطفاً کپچای ظاهرشده در مرورگر را حل کنید تا استخراج ادامه یابد.',
            )

            while not self.stop_requested:
              time.sleep(1.5)
              cur_title = page.title().lower()
              cur_url = page.url.lower()
              still_blocked = (
                  'captcha' in cur_url
                  or '403' in cur_title
                  or 'forbidden' in cur_title
                  or 'تایید هویت' in cur_title
              )
              if not still_blocked:
                self.log('✅ کپچا حل شد! ادامه استخراج...\n')
                time.sleep(1)
                break
        except Exception:
          pass

      def safe_fetch_json(self, page, url):
        js_fetch = """async (target_url) => {
                try {
                    const res = await fetch(target_url, {
                        headers: {
                            'accept': 'application/json, text/plain, */*',
                            'accept-language': 'fa-IR,fa;q=0.9,en;q=0.8'
                        }
                    });
                    if (res.status === 403 || res.status === 429) {
                        return { status: res.status, data: null };
                    }
                    if (!res.ok) return { status: res.status, data: null };
                    const data = await res.json();
                    return { status: 200, data: data };
                } catch(e) {
                    return { status: 500, data: null };
                }
            }"""

        if 'torob.com' not in page.url or 'api.torob.com' in page.url:
          page.goto(
              'https://torob.com/', wait_until='domcontentloaded', timeout=30000
          )

        res = page.evaluate(js_fetch, url)

        if res and res.get('status') in [403, 429]:
          page.goto(
              'https://torob.com/', wait_until='domcontentloaded', timeout=30000
          )
          self.handle_captcha_if_present(page)
          res = page.evaluate(js_fetch, url)

        return res.get('data') if res else None

      def start_scraping_thread(self):
        if not self.selected_file_path:
          messagebox.showwarning('خطا', 'لطفاً ابتدا فایل اکسل را انتخاب کنید.')
          return

        try:
          start_row = int(self.entry_start_row.get())
          end_row = int(self.entry_end_row.get())
          #max_results = int(self.entry_max_results.get())

      
        except ValueError:
          messagebox.showerror('خطا', 'لطفاً مقادیر معتبر وارد کنید.')
          return

        self.is_running = True
        self.stop_requested = False
        self.btn_start.configure(state='disabled')
        self.btn_stop.configure(state='normal')
        self.btn_select_file.configure(state='disabled')

        is_headless = bool(self.switch_headless.get())

        threading.Thread(
            target=self.run_scraper,
            args=(start_row, end_row, max_results, is_headless),
            daemon=True,
        ).start()
  
      def wait_for_unblock(self):
          self.log("صبر یک ساعته به دلیل مسدودیت")
          elapsed = 0
          while elapsed < 3600 and not self.stop_requested:
              time.sleep(5)
              elapsed += 5

      def run_scraper(self, start_row, end_row, max_results, is_headless):
        try:
          input_df = pd.read_excel(self.selected_file_path)
          total_rows = len(input_df)
          sliced_df = input_df.iloc[start_row - 1 : end_row]

          now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
          base_filename = f'Torob_Report_Rows_{start_row}_to_{end_row}_{now}'
          excel_filename = f'{base_filename}.xlsx'
          json_filename = f'{base_filename}.json'

          wb = openpyxl.Workbook()
          ws = wb.active
          ws.title = 'گزارش ترب'

          headers = [
              'کد کالا',
              'نام کالا',
              'نام در ترب',
              'فروشگاه',
              'امتیاز و سابقه',
              'ضمانت ترب',
              'قیمت',
              'تغییر قیمت',
              'لینک',
              'عکس',
          ]
          ws.append(headers)

          header_fill = PatternFill(
              start_color='1F4E78', end_color='1F4E78', fill_type='solid'
          )
          header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
          for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

          # رنگ‌های متناوب سفید و کرمی برای تمایز محصولات ورودی
          fill_white = PatternFill(
              start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'
          )
          fill_cream = PatternFill(
              start_color='FFF9E6', end_color='FFF9E6', fill_type='solid'
          )

          json_export_list = []
          use_cream = False  # پرچم تغییر رنگ بین محصولات

          self.log(
              f'\n🚀 شروع استخراج سطرهای {start_row} تا'
              f' {min(end_row, total_rows)}...'
          )

          with sync_playwright() as p:
            launch_args = [
                '--disable-blink-features=AutomationControlled',
                '--no-sandbox',
            ]
            try:
              browser = p.chromium.launch(
                  headless=is_headless, args=launch_args, channel='msedge'
              )
            except Exception:
              browser = p.chromium.launch(headless=is_headless, args=launch_args)

            context = browser.new_context(
                user_agent=(
                    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    ' (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
                ),
                locale='fa-IR',
            )
            page = context.new_page()

            page.goto(
                'https://torob.com/', wait_until='domcontentloaded', timeout=60000
            )
            time.sleep(2)
            self.handle_captcha_if_present(page)

            for idx, (_, row) in enumerate(sliced_df.iterrows()):
              if self.stop_requested:
                break

              product_code = row.get('کد', '-')
              product_name = str(row.get('نام', '')).strip()

              if not product_name or pd.isna(product_name):
                continue

              # انتخاب رنگ برای این محصول ورودی
              current_fill = fill_cream if use_cream else fill_white
              use_cream = not use_cream  # تغییر رنگ برای محصول بعدی

              self.log(f'\n🔍 [سطر {start_row + idx}] جستجو: {product_name}')
              while True:
                  time.sleep(random.uniform(2, 5))

                  search_api_url = f'https://api.torob.com/v4/base-product/search/?q={urllib.parse.quote(product_name)}&page=0&size=24'
                  search_data = self.safe_fetch_json(page, search_api_url)

                  results = (
                      search_data.get('results', [])
                      if isinstance(search_data, dict)
                      else []
                  )
              
                  if not results:
                      self.log(f'❌ یافت نشد: {product_name}')
                      self.wait_for_unblock()
                  else:
                      break
                        
                for cell in ws[ws.max_row]:
                  cell.fill = current_fill
                continue

              target_results = results[:max_results]

              for top_product in target_results:
                if self.stop_requested:
                  break

                prk = top_product.get('random_key') or top_product.get('prk')
                image_url = top_product.get('image_url', '-')

                if prk:
                  sellers_url = (
                      f'https://api.torob.com/v4/base-product/sellers/?prk={prk}'
                  )
                  sellers_data = self.safe_fetch_json(page, sellers_url)

                  sellers_list = []
                  if sellers_data and isinstance(sellers_data, list):
                    sellers_list = sellers_data
                  elif sellers_data and isinstance(sellers_data, dict):
                    sellers_list = sellers_data.get('results', [])

                  if not sellers_list:
                    details_url = (
                        f'https://api.torob.com/v4/base-product/details/?prk={prk}'
                    )
                    details_data = self.safe_fetch_json(page, details_url)

                    if details_data and isinstance(details_data, dict):
                      sellers_list = details_data.get('products_info', {}).get(
                          'result', []
                      )

                  for seller in sellers_list:
                    torob_title = seller.get('name1') or top_product.get(
                        'name1', '-'
                    )

                    guarantee_info = seller.get('guarantee_info') or {}
                    torob_guarantee = (
                        'دارد'
                        if guarantee_info.get('status') == 'enabled'
                        else 'ندارد'
                    )

                    shop_name = seller.get('shop_name') or '-'

                    score_info = seller.get('score_info') or {}
                    score_and_history = (
                        score_info.get('score_text')
                        or seller.get('shop_score_text')
                        or '-'
                    )

                    price_change = seller.get('last_price_change_date') or '-'
                    price = (
                        seller.get('price_text')
                        or seller.get('price_string')
                        or '-'
                    )
                    link = seller.get('page_url') or '-'

                    row_data = [
                        product_code,
                        product_name,
                        torob_title,
                        shop_name,
                        score_and_history,
                        torob_guarantee,
                        price,
                        price_change,
                        link,
                        image_url,
                    ]
                    ws.append(row_data)

                    # اعمال رنگ متناوب محصول روی تمامی سلول‌های سطر جدید
                    for cell in ws[ws.max_row]:
                      cell.fill = current_fill

                    json_export_list.append({
                        'product_code': product_code,
                        'product_name': product_name,
                        'torob_title': torob_title,
                        'shop_name': shop_name,
                        'score_and_history': score_and_history,
                        'torob_guarantee': torob_guarantee,
                        'price': price,
                        'price_change': price_change,
                        'link': link,
                        'image_url': image_url,
                        'prk': prk,
                        'raw_seller_data': seller,
                    })

              self.log(f'✅ موفق: {product_name}')

            wb.save(excel_filename)

            with open(json_filename, 'w', encoding='utf-8') as f:
              json.dump(json_export_list, f, ensure_ascii=False, indent=2)

            if TorobChartModule:
              chart_module = TorobChartModule(excel_filename)
              chart_module.generate_charts()
              self.log('📊 نمودارهای تحلیلی کلی با موفقیت به اکسل اضافه شدند.')


            if TorobChartModule:
              self.log('\n📸 در حال باز کردن صفحات و ثبت اسکرین‌شات نمودارها...')
              chart_module = TorobChartModule(excel_filename)

              product_prks_map = {}
              for item in json_export_list:
                p_name = item.get('product_name')
                prk_val = item.get('prk')
                if p_name and prk_val:
                  if p_name not in product_prks_map:
                    product_prks_map[p_name] = []
                  if prk_val not in product_prks_map[p_name]:
                    product_prks_map[p_name].append(prk_val)

              image_row = 2
              for idx, (prod_name, prk_list) in enumerate(
                  list(product_prks_map.items())[:10]
              ):
                if self.stop_requested:
                  break

                selected_prk = prk_list[1] if len(prk_list) > 1 else prk_list[0]
                result_tag = (
                    'نتیجه دوم (واقعی)' if len(prk_list) > 1 else 'تنها نتیجه'
                )

                self.log(
                    f'📸 ثبت اسکرین‌شات نمودار ({result_tag}): {prod_name}'
                )
                prod_url = f'https://torob.com/p/{selected_prk}/'

                try:
                  page.goto(prod_url, wait_until='domcontentloaded', timeout=30000)
                  time.sleep(2)

                  self.handle_captcha_if_present(page)

                  # ۱. یافتن دکمه بازکردن نمودار با استفاده از aria-label و الگوی کلاس
                  chart_btn = page.locator(
                      'button[aria-label="لیست تغییرات قیمت"],'
                      ' button[class*="chartOpenButton"]'
                  ).first

                  if chart_btn.count() > 0:
                    chart_btn.scroll_into_view_if_needed()
                    time.sleep(0.5)
                    chart_btn.click()
                    time.sleep(2)  # زمان برای باز شدن مدال و رندر محورهای قیمت و تاریخ

                    img_filename = f'temp_chart_{idx}.png'

                    # ۲. گرفتن اسکرین‌شات از کادر مدال بازشده
                    modal_element = page.locator('div[role="dialog"]').first
                    if modal_element.count() > 0 and modal_element.is_visible():
                      modal_element.screenshot(path=img_filename)
                    else:
                      page.screenshot(path=img_filename)

                    chart_module.add_chart_image(prod_name, img_filename, image_row)
                    image_row += 15

                    # ۳. بستن مدال با دکمه Esc
                    page.keyboard.press('Escape')
                    time.sleep(0.8)

                    if os.path.exists(img_filename):
                      os.remove(img_filename)
                  else:
                    self.log(f'⚠️ دکمه نمودار قیمت برای {prod_name} یافت نشد.')

                except Exception as err:
                  self.log(f'❌ خطا در اسکرین‌شات {prod_name}: {err}')

                time.sleep(1.5)

              self.log('✅ تصاویر نمودار با موفقیت ثبت و درج شدند.')

            browser.close()

          self.log(
              f'\n💾 فایل اکسل: {excel_filename}\n💾 فایل JSON:'
              f' {json_filename}'
          )
          messagebox.showinfo(
              'پایان عملیات',
              f'استخراج کامل شد.\n\nاکسل: {excel_filename}\nجیسون: {json_filename}',
          )

        except Exception as e:
          self.log(f'\n❌ خطا: {e}')
          messagebox.showerror('خطا', f'خطایی رخ داد: {e}')

        finally:
          self.is_running = False
          self.btn_start.configure(state='normal')
          self.btn_stop.configure(state='disabled')
          self.btn_select_file.configure(state='normal')


if __name__ == '__main__':
  app = TorobScraperApp()
  app.mainloop()

