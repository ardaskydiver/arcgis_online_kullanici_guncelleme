import pandas as pd
from arcgis.gis import GIS
import requests

# Excel dosyasini oku/Read Excel file
df = pd.read_excel(r"D:\\online_kullanici_degistirme\\kullanici_bilgileri.xlsx", dtype={'kullanici_adi':str, 'sifre':str , 'yeni_sifre':str})

# Her bir kullanici icin islemleri gerceklestir/Perform operations for each user
for index, row in df.iterrows():
    kullanici_adi = row['kullanici_adi']
    sifre = row['sifre']
    yeni_sifre = row['yeni_sifre']

    # ArcGIS Online'a baglan/Connect ArcGIS Online
    gis = GIS("https://www.arcgis.com", kullanici_adi, sifre)
    print(f"Basarili bir sekilde giris yapildi: {gis.properties.user.username}")

    # Hesaptaki tum ogeleri al ve sil/Delete all items on username
    items = gis.content.search(query=f"owner:{kullanici_adi}", max_items=10000)
    for item in items:
        item.delete()
        print(f"'{item.title}' basariyla silindi.")
    print("Tum icerikler silindi.")

    # uyesi oldugunuz tum gruplari alin ve gruptan cik/Delete group from username 
    gruplar = gis.groups.search(query="", max_groups=500)
    for grup in gruplar:
        if grup.owner != gis.users.me.username:
            try:
                grup.leave()
                print(f"'{grup.title}' grubundan basariyla cikis yapildi.")
            except Exception as e:
                pass
                #print(f"'{grup.title}' grubundan cikis yapilirken bir hata olustu: {e}")
    print("Tum gruplardan cikis islemi tamamlandi.")

    # sifre degistirme islemi/Change password
    url = f"https://www.arcgis.com/sharing/rest/community/users/{kullanici_adi}/update"
    params = {
        'f': 'json',
        'token': gis._con.token,
        'password': yeni_sifre,
        'currentPassword': sifre
    }
    response = requests.post(url, data=params)
    if response.status_code == 200:
        print("sifre basariyla degistirildi.")
    else:
        print(f"sifre degistirilemedi. Hata mesaji: {response.json()}")

print("Tum kullanicilar icin islemler tamamlandi.")
