from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import strftime
import random
import time
import ctypes
import os


if os.name == 'nt':
    ctypes.windll.kernel32.SetConsoleTitleW("Automedic BETA by kakegemesh")

os.system('cls' if os.name == 'nt' else 'clear')

filePath = (os.path.abspath(os.path.dirname(os.path.realpath(__file__))))
browserPath = filePath + "\chromedriver.exe"

jam = strftime("%H:%M:%S")
h = int(strftime("%H"))
m = int(strftime("%M"))

print("  ___        _         ___  ___         _ _      ")
print(" / _ \      | |        |  \/  |        | (_)     ")
print("/ /_\ \_   _| |_ ___   | .  . | ___  __| |_  ___ ")
print("|  _  | | | | __/ _ \  | |\/| |/ _ \/ _` | |/ __|")
print("| | | | |_| | || (_) | | |  | |  __/ (_| | | (__ ")
print("\_| |_/\__,_|\__\___/  \_|  |_/\___|\__,_|_|\___|")
                                                 
                                                 
print("\n")
print("\n")
print('Sekarang jam ' + jam )
if h < 9 and m <= 15:
    print("Belum terlambat buat ngisi daily medical check up")
else:
    print("Sebenernya udah terlambat sih buat ngisi medical check up")
print("Automedic akan mengisi daily medical kamu")
print("Isi nomer NPK karyawannya ya")

npk = "2110210"

list_npk = ["2110210", "2141092", "2181001", "999999"]
list_nama = ["Dedy Supriyanto", "Izzi Rohmatulloh", "Gilang Fajar Setiawan", "Mukidi"]

for i in list_npk:
    if i == npk:
        y = list_npk.index(npk)
        nama = list_nama[y]
    #elif i != npk: 
        # print("NPK Tidak Terdaftar")
        #keluarapk = input("Tutup aplikasi? (Y) = ")
            #if keluarapk == "Y" or "y":
            #exit()

x = "36."
y = random.randint(4,7)
suhu = (x + str(y))
print('Suhu kamu hari ini ' + suhu)

print("OK " + nama + " tunggu sebentar ya...")
time.sleep(2)
print("Starting Automedic...")

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation']);

if os.name == 'nt':
    filePath = (os.path.abspath(os.path.dirname(os.path.realpath(__file__))))
    browserPath = filePath + "\chromedriver.exe"
    browser = webdriver.Chrome(executable_path=browserPath)
else:
    browser = webdriver.Safari()

# Membuka Website Medical Form
browser.get('https://edmcu.tacindonesia.id/admin/index.php')

time.sleep(1)

os.system('cls' if os.name == 'nt' else 'clear')
print("Sedang Dalam Proses Pengisian Form Daily Medical...")
time.sleep(1)

# Mengisi NPK dan Nama
tbox = browser.find_elements_by_class_name("form-control")
#tb2 =  browser.find_elements_by_xpath("office-form-question-textbox")

tbox[0].send_keys(npk)
tbox[0].send_keys(Keys.SPACE)
time.sleep(2)

kesehatan = browser.find_element(By.XPATH, "//*[@id='kondisiKesehatan2']") # Pilih Kondisi Kesehatan
kesehatan.click()
time.sleep(0.5)
omicron = browser.find_element(By.XPATH, "//*[@value='Tidak ada Gejala tersebut diatas']") # Gejala Omicron
omicron.click()
time.sleep(0.5)

hadir = browser.find_element(By.XPATH, "//*[@id='statusKehadiran']") # Pilih Status Kehadiran
hadir.click()
time.sleep(0.5)
bekerja = browser.find_element(By.XPATH, "//*[@value='Masuk Bekerja']") # Masuk Bekerja
bekerja.click()
time.sleep(0.5)

shift = browser.find_element(By.XPATH, "//*[@id='shift']") # Pilih Shift
shift.click()
time.sleep(0.5)
nonshift = browser.find_element(By.XPATH, "//*[@value='Non Shift']") # Non Shift 
nonshift.click()
time.sleep(0.5)

puasa = browser.find_element(By.XPATH, "//*[@id='puasa']") # Pilihan Puasa
puasa.click()
time.sleep(0.5)
tidakPuasa = browser.find_element(By.XPATH, "//*[@value='Tidak']") # Tidak Puasa
tidakPuasa.click()
time.sleep(0.5)

kesehatan2 = browser.find_element(By.XPATH, "//*[@id='kondisiKesehatan']") # Pilihan Kesehatan
kesehatan2.click()
time.sleep(0.5)
kondisiSehat = browser.find_element(By.XPATH, "//*[@value='SEHAT']") # Sehat Alhamdulillah
kondisiSehat.click()
time.sleep(0.5)

saveButton = browser.find_element(By.XPATH, "//*[@id='tombolSave']") # Pilihan Kesehatan
saveButton.click()
time.sleep(0.5)
