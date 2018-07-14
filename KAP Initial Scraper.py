import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import xlsxwriter

str_ana_sayfa = 'https://www.kap.org.tr'
r = requests.get('https://www.kap.org.tr/tr/bist-sirketler')
s = BeautifulSoup(r.content, 'html.parser')

vFind_all1 = s.find_all(class_='column-type7 wmargin')
vSirketAdi = s.find_all(class_='comp-cell _14 vtable')
vFirmaKod = s.find_all(class_='comp-cell _04 vtable')
vSehir = s.find_all(class_='comp-cell _12 vtable')
vBagDeg = s.find_all(class_='comp-cell _11 vtable')
vLink = s.select('div.comp-cell._04.vtable')

firma_isim_list = []
firma_kod_list = []
firma_seh_list = []
firma_bd_list = []
firma_link_list = []
firma_link_list_ozet = []

i = 0
ii = 0
iii = 0
ix = 0
ixi = 0
ixi2 = 0

while i < len(vSirketAdi):
    ham_adlar = vSirketAdi[i].get_text()
    ap = ham_adlar.replace("\n", "")
    firma_isim_list.append(ap)
    i = i + 1

while ii < len(vFirmaKod):
    ham_kodlar = vFirmaKod[ii].get_text()
    ap = ham_kodlar.replace("\n", "")
    firma_kod_list.append(ap)
    ii = ii + 1

while iii < len(vSehir):
    ham_kodlar = vSehir[iii].get_text()
    ap = ham_kodlar.replace("\n", "")
    firma_seh_list.append(ap)
    iii = iii + 1

while ix < len(vBagDeg):
    ham_kodlar = vBagDeg[ix].get_text()
    ap = ham_kodlar.replace("\n", "")
    firma_bd_list.append(ap)
    ix = ix + 1

while ixi < len(vLink):
    vLink_t = vLink[ixi]
    vLink_t2 = vLink_t.find('a')
    vLink_href = str(vLink_t2['href']).replace('ozet', 'genel')
    vAppend_link = str_ana_sayfa + vLink_href
    firma_link_list.append(vAppend_link)
    ixi = ixi + 1

while ixi2 < len(vLink):
    vLink_t = vLink[ixi2]
    vLink_t2 = vLink_t.find('a')
    vLink_href = str(vLink_t2['href'])
    vAppend_link = str_ana_sayfa + vLink_href
    firma_link_list_ozet.append(vAppend_link)
    ixi2 = ixi2 + 1

genel_data = pd.DataFrame({'Ad': firma_isim_list,
                           'Kod': firma_kod_list,
                           'Sehir': firma_seh_list,
                           'Bag_Denetim': firma_bd_list,
                           'Link': firma_link_list,
                           'Özet Link': firma_link_list_ozet,
                           })

genel_data.to_excel('BIST_Sirketleri.xlsx', sheet_name='Sheet1')

ozet_son_liste = []

def OzetBilgi(ls, kod, timeout):
    "Firma özet bilgilerini çekip ozet_bilgiler listesini uygun şekilde doldurur"
    it = 0

    while it < len(ls):
        kk = kod[it]
        ozet_son_liste.append(kk)

        ozet_veri_liste = []
        ozet_baslik_list = []
        ozet_dict = {'Adres': '',
                     'Eposta': '',
                     'Web': '',
                     'Sure': '',
                     'BagDeg': '',
                     'Endeks': '',
                     'Sektor': '',
                     'Pazar': '',
                     'ABCD': '',
                     }

        str_ls = ls[it]
        req = requests.get(str_ls)
        sp = BeautifulSoup(req.content, 'html.parser')
        vfFind_all = sp.find_all(class_='comp-cell-row-div vtable infoColumn backgroundThemeForValue')
        vFindall_basliklar = sp.find_all(class_='comp-cell-row-div vtable infoColumn backgroundThemeForTitle')
        now = time.strftime("%d-%m-%Y %H:%M:%S")

        with open("Ozet_Log.txt", "a") as log_file:
            log_file.write(now + "-")
            log_file.write(kk)
            log_file.write("\n")

        for l in vfFind_all:
            zz = l.get_text()
            zzz = zz.replace("\n", "").strip()
            ozet_veri_liste.append(zzz)
        for b in vFindall_basliklar:
            zz2 = b.get_text()
            zzz2 = zz2.replace("\n", "").strip()
            ozet_baslik_list.append(zzz2)

        ima = ''
        ipo = ''
        iint = ''
        isu = ''
        ibd = ''
        ien = ''
        isek = ''
        ipaz = ''
        iab = ''

        if ozet_baslik_list.count('Merkez Adresi') > 0:
            ima = ozet_baslik_list.index('Merkez Adresi')
            ozet_dict['Adres'] = ozet_veri_liste[ima]
        if ozet_baslik_list.count('Elektronik Posta Adresi') > 0:
            ipo = ozet_baslik_list.index('Elektronik Posta Adresi')
            ozet_dict['Eposta'] = ozet_veri_liste[ipo]
        if ozet_baslik_list.count('İnternet Adresi') > 0:
            iint = ozet_baslik_list.index('İnternet Adresi')
            ozet_dict['Web'] = ozet_veri_liste[iint]
        if ozet_baslik_list.count('Şirketin Süresi') > 0:
            isu = ozet_baslik_list.index('Şirketin Süresi')
            ozet_dict['Sure'] = ozet_veri_liste[isu]
        if ozet_baslik_list.count('Bağımsız Denetim Kuruluşu') > 0:
            ibd = ozet_baslik_list.index('Bağımsız Denetim Kuruluşu')
            ozet_dict['BagDeg'] = ozet_veri_liste[ibd]
        if ozet_baslik_list.count('Şirketin Dahil Olduğu Endeksler') > 0:
            ien = ozet_baslik_list.index('Şirketin Dahil Olduğu Endeksler')
            ozet_dict['Endeks'] = ozet_veri_liste[ien]
        if ozet_baslik_list.count('Şirketin Sektörü') > 0:
            isek = ozet_baslik_list.index('Şirketin Sektörü')
            ozet_dict['Sektor'] = ozet_veri_liste[isek]
        if ozet_baslik_list.count('Sermaye Piyasası Aracının İşlem Gördüğü Pazar') > 0:
            ipaz = ozet_baslik_list.index('Sermaye Piyasası Aracının İşlem Gördüğü Pazar')
            ozet_dict['Pazar'] = ozet_veri_liste[ipaz]
        if ozet_baslik_list.count('ABCD Grubu') > 0:
            iab = ozet_baslik_list.index('ABCD Grubu')
            ozet_dict['ABCD'] = ozet_veri_liste[iab]
        for tt in ozet_dict.values():
            ozet_son_liste.append(tt)

        it = it + 1
        time.sleep(timeout)
    return ozet_son_liste

OzetBilgi(firma_link_list_ozet, firma_kod_list,5)

ozet_bilgiler_kod = []
ozet_bilgiler_adres = []
ozet_bilgiler_mail = []
ozet_bilgiler_web = []
ozet_bilgiler_sure = []
ozet_bilgiler_bagden = []
ozet_bilgiler_endks = []
ozet_bilgiler_sekt = []
ozet_bilgiler_pazar = []
ozet_bilgiler_abcd = []

ozi = 0
for ozz in ozet_son_liste:
    if ozi % 10 == 0:
        ozet_bilgiler_kod.append(ozz)
    elif ozi % 10 == 1:
        ozet_bilgiler_adres.append(ozz)
    elif ozi % 10 == 2:
        ozet_bilgiler_mail.append(ozz)
    elif ozi % 10 == 3:
        ozet_bilgiler_web.append(ozz)
    elif ozi % 10 == 4:
        ozet_bilgiler_sure.append(ozz)
    elif ozi % 10 == 5:
        ozet_bilgiler_bagden.append(ozz)
    elif ozi % 10 == 6:
        ozet_bilgiler_endks.append(ozz)
    elif ozi % 10 == 7:
        ozet_bilgiler_sekt.append(ozz)
    elif ozi % 10 == 8:
        ozet_bilgiler_pazar.append(ozz)
    elif ozi % 10 == 9:
        ozet_bilgiler_abcd.append(ozz)
    ozi = ozi + 1

ozet_bilgiler_pd = pd.DataFrame({'Kod': ozet_bilgiler_kod,
                                 'Adres': ozet_bilgiler_adres,
                                 'Email': ozet_bilgiler_mail,
                                 'Web Adres': ozet_bilgiler_web,
                                 'Süre': ozet_bilgiler_sure,
                                 'Bağımsız Denetim': ozet_bilgiler_bagden,
                                 'Endeksler': ozet_bilgiler_endks,
                                 'Sektör': ozet_bilgiler_sekt,
                                 'Pazar': ozet_bilgiler_pazar,
                                 'ABCD': ozet_bilgiler_abcd,
                                 })

ozet_bilgiler_pd.to_excel('Firma_Ozet_Bilgileri.xlsx', sheet_name='Ozet')


def GenelBilgiler(genel_linkler, kodlar, timeout):
    vit = 0
    yipd = pd.DataFrame()
    ykpd = pd.DataFrame()
    ysspd = pd.DataFrame()
    ortpd = pd.DataFrame()
    dolpd = pd.DataFrame()
    dppd = pd.DataFrame()
    istpd = pd.DataFrame()

    while vit < len(genel_linkler):
        firma_kod = kodlar[vit]

        str_gen_link = genel_linkler[vit]
        r_genel = requests.get(str_gen_link)
        s_genel = BeautifulSoup(r_genel.content, 'html.parser')
        vFind_all1 = s_genel.find_all(class_=re.compile("comp-cell-row-div vtable infoColumn*"))
        vFind_all2 = s_genel.find_all(True, {'class': [re.compile("comp-cell-row-div vtable infoColumn*"), re.compile("type-normal vcell exportTitle*"), 'column-type3 exportDiv']})

        with open("Log.txt", "a") as log_file:
            now2 = time.strftime("%d-%m-%Y %H:%M:%S")
            log_file.write(now2 + "-")
            log_file.write(firma_kod)
            log_file.write("\n")

        ham_liste = []
        ham_liste2 = []

        for x in vFind_all1:
            vv = x.get_text()
            vvv = vv.replace("\n", "").strip()
            ham_liste.append(vvv)

        for y in vFind_all2:
            vv2 = y.get_text()
            vvv2 = vv2.replace("\n", " ").strip()
            ham_liste2.append(vvv2)

        def Yat_ilis(lis2, kod):

            """Yatırımcı ilişkilerinden sorumlu kişilerin Kod-Ad-Görev formatında pandas_dataFrame şeklinde listeler"""

            yat_ilis_ad = []
            yat_ilis_gorev = []

            yat_ilis_start = lis2.index('Yatırımcı İlişkileri Bölümü veya Bağlantı Kurulacak Şirket Yetkilileri') + 8
            yat_ilis_pre_list = lis2[yat_ilis_start:]

            if yat_ilis_pre_list[0] != 'Sermaye Piyasası Aracının İşlem Gördüğü Pazar':

                if yat_ilis_pre_list.count('Merkez Dışı Örgütleri (Şube, İrtibat Bürosu)') > 0:
                    yat_ilis_end = yat_ilis_pre_list.index('Merkez Dışı Örgütleri (Şube, İrtibat Bürosu)')
                else:
                    yat_ilis_end = yat_ilis_pre_list.index('Şirketin Faaliyet Konusu')

                yat_ilis_list = yat_ilis_pre_list[:yat_ilis_end]

                i = 0
                for y in yat_ilis_list:
                    if i % 7 == 0:
                        yat_ilis_ad.append(y)
                    i = i + 1

                ii = 0
                for yy in yat_ilis_list:
                    if (ii - 1) % 7 == 0:
                        yat_ilis_gorev.append(yy)
                    ii = ii + 1

            else:
                yat_ilis_ad.append(' ')
                yat_ilis_gorev.append(' ')

            yat_ilis_pd = pd.DataFrame({'Kod': kod,
                                        'Yat_ilis_ad': yat_ilis_ad,
                                        'Yat_ilis_gorev': yat_ilis_gorev,
                                        })

            return yat_ilis_pd

        def Yonetim_Kurulu(lis, kod):
            """Yönetim kurulu bilgilerini Kod-Ad-tüzel-cins-görev-meslek-son5-ortaklıkdışı-sermaye payı- pay grubu- bağımsız üye-komiteler formatında pandas_dataFrame şeklinde listeler"""

            yon_kur_ad = []
            yon_kur_tuz_ad = []
            yon_kur_cins = []
            yon_kur_gorev = []
            yon_kur_meslek = []
            yon_kur_son5 = []
            yon_kur_disgorev = []
            yon_kur_serpay = []
            yon_kur_temsilpay = []
            yon_kur_bagim = []
            yon_kur_komite = []

            if lis.count('Tüzel Kişi Üye Adına Hareket Eden Kişi') > 0:

                yon_kur_start = lis.index('Tüzel Kişi Üye Adına Hareket Eden Kişi') + 10

                if lis.count('Son Durum İtibariyle Ortaklık Dışında Aldığı Görevler') > 0:
                    yon_kur_end = lis.index('Son Durum İtibariyle Ortaklık Dışında Aldığı Görevler') - 4
                else:
                    yon_kur_end = lis.index('Son Durum itibariyle Ortaklık Dışında Aldığı Görevler') - 4

                yon_kur_list = lis[yon_kur_start:yon_kur_end]

                yk_i = 0
                for yk in yon_kur_list:
                    if yk_i % 11 == 0:
                        yon_kur_ad.append(yk)
                    elif yk_i % 11 == 1:
                        yon_kur_tuz_ad.append(yk)
                    elif yk_i % 11 == 2:
                        yon_kur_cins.append(yk)
                    elif yk_i % 11 == 3:
                        yon_kur_gorev.append(yk)
                    elif yk_i % 11 == 4:
                        yon_kur_meslek.append(yk)
                    elif yk_i % 11 == 5:
                        yon_kur_son5.append(yk)
                    elif yk_i % 11 == 6:
                        yon_kur_disgorev.append(yk)
                    elif yk_i % 11 == 7:
                        yon_kur_serpay.append(yk)
                    elif yk_i % 11 == 8:
                        yon_kur_temsilpay.append(yk)
                    elif yk_i % 11 == 9:
                        yon_kur_bagim.append(yk)
                    elif yk_i % 11 == 10:
                        yon_kur_komite.append(yk)

                    yk_i = yk_i + 1
            else:
                yon_kur_ad.append(' ')
                yon_kur_tuz_ad.append(' ')
                yon_kur_cins.append(' ')
                yon_kur_gorev.append(' ')
                yon_kur_meslek.append(' ')
                yon_kur_son5.append(' ')
                yon_kur_disgorev.append(' ')
                yon_kur_serpay.append(' ')
                yon_kur_temsilpay.append(' ')
                yon_kur_bagim.append(' ')
                yon_kur_komite.append(' ')

            yon_kur_pd = pd.DataFrame({'Kod': kod,
                                       'YK_ad': yon_kur_ad,
                                       'Tüzel Kişi Adına': yon_kur_tuz_ad,
                                       'Cinsiyet': yon_kur_cins,
                                       'Görev': yon_kur_gorev,
                                       'Meslek': yon_kur_meslek,
                                       'Son 5 yılda': yon_kur_son5,
                                       'Ortaklık dışı görev': yon_kur_disgorev,
                                       'Sermaye Payı': yon_kur_serpay,
                                       'Pay Grubu': yon_kur_temsilpay,
                                       'Bağımsız Mı': yon_kur_bagim,
                                       'Komiteler': yon_kur_komite,
                                       })

            return yon_kur_pd

        def Yonetim_Soz_Sahibi(lis, kod):
            """Yatırımcı ilişkilerinden sorumlu kişilerin Kod-Ad-Görev-meslek-son 5 yılda-ortaklık dışı görev formatında pandas_dataFrame şeklinde listeler"""

            soz_sah_ad = []
            soz_sah_gorev = []
            soz_sah_meslek = []
            soz_sah_son5 = []
            soz_sah_disgorev = []

            if lis.count('Son Durum İtibariyle Ortaklık Dışında Aldığı Görevler') > 0:

                soz_sah_start = lis.index('Son Durum İtibariyle Ortaklık Dışında Aldığı Görevler') + 1
                soz_sah_sliced_list = lis[soz_sah_start:]
                soz_sah_ort_start = len(soz_sah_sliced_list)
                soz_sah_pay_start = len(soz_sah_sliced_list)
                soz_sah_ad_start = len(soz_sah_sliced_list)
                soz_sah_borsa_kodu_start = len(soz_sah_sliced_list)
                soz_sah_tic_unvan_start = len(soz_sah_sliced_list)

                if soz_sah_sliced_list.count('Ortağın Adı-Soyadı/Ticaret Unvanı') > 0:
                    soz_sah_ort_start = soz_sah_sliced_list.index('Ortağın Adı-Soyadı/Ticaret Unvanı')
                if soz_sah_sliced_list.count('Pay Grubu') > 0:
                    soz_sah_pay_start = soz_sah_sliced_list.index('Pay Grubu')
                if soz_sah_sliced_list.count('Adı-Soyadı') > 0:
                    soz_sah_ad_start = soz_sah_sliced_list.index('Adı-Soyadı')
                if soz_sah_sliced_list.count('Borsa Kodu') > 0:
                    soz_sah_borsa_kodu_start = soz_sah_sliced_list.index('Borsa Kodu')
                if soz_sah_sliced_list.count('Ticaret Unvanı') > 0:
                    soz_sah_tic_unvan_start = soz_sah_sliced_list.index('Ticaret Unvanı')

                soz_sah_end = min(soz_sah_ort_start, soz_sah_pay_start, soz_sah_ad_start, soz_sah_borsa_kodu_start,soz_sah_tic_unvan_start)
                soz_sah_list = soz_sah_sliced_list[:soz_sah_end]

                sozi = 0
                for soz in soz_sah_list:
                    if sozi % 5 == 0:
                        soz_sah_ad.append(soz)
                    elif sozi % 5 == 1:
                        soz_sah_gorev.append(soz)
                    elif sozi % 5 == 2:
                        soz_sah_meslek.append(soz)
                    elif sozi % 5 == 3:
                        soz_sah_son5.append(soz)
                    elif sozi % 5 == 4:
                        soz_sah_disgorev.append(soz)
                    sozi = sozi + 1
            else:

                soz_sah_ad.append(' ')
                soz_sah_gorev.append(' ')
                soz_sah_meslek.append(' ')
                soz_sah_son5.append(' ')
                soz_sah_disgorev.append(' ')

            soz_sah_pd = pd.DataFrame({'Kod': kod,
                                       'Soz_sah_ad': soz_sah_ad,
                                       'Görev': soz_sah_gorev,
                                       'Meslek': soz_sah_meslek,
                                       'Son 5 yılda': soz_sah_son5,
                                       'Ortaklık dışı görev': soz_sah_disgorev,
                                       })

            return soz_sah_pd

        def Ortaklar(lis2, kod):
            """Firma ortaklarını kod-ad-pay TL-Pay oranı-oy oranı formatında pandasDataFrame listeler"""

            ortak_ad_unvan = []
            ortak_pay = []
            ortak_payoran = []
            ortak_pb = []
            ortak_oyoran = []

            if lis2.count('Ortaklık Yapısı') > 0:
                ortak_start = lis2.index('Ortaklık Yapısı') + 5
                ortak_mini_list = lis2[ortak_start:]

                if ortak_mini_list[0] != 'Bilgi Mevcut Değil' and ortak_mini_list[0] != 'Ödenmiş/Çıkarılmış Sermayesi':
                    ortak_end = ortak_mini_list.index('TOPLAM') + 4
                    ortak_list = ortak_mini_list[:ortak_end]

                    ort = 0
                    for ori in ortak_list:

                        if ort % 4 == 0:
                            ortak_ad_unvan.append(ori)
                        elif ort % 4 == 1:
                            ortak_pay.append(ori)
                        elif ort % 4 == 2:
                            ortak_pb.append(ori)
                        elif ort % 4 == 3:
                            ortak_payoran.append(ori)
                            ortak_oyoran.append(ori)
                        ort = ort + 1
                else:
                    ortak_ad_unvan.append(' ')
                    ortak_pay.append(' ')
                    ortak_payoran.append(' ')
                    ortak_pb.append(' ')
                    ortak_oyoran.append(' ')

            else:
                ortak_start = lis2.index(
                    'Sermayede Doğrudan %5 veya Daha Fazla Paya veya Oy Hakkına Sahip Gerçek ve Tüzel Kişiler') + 5
                ortak_mini_list = lis2[ortak_start:]

                if ortak_mini_list.count('TOPLAM') > 0:

                    ortak_end = ortak_mini_list.index('TOPLAM') + 4
                    ortak_list = ortak_mini_list[:ortak_end]
                    ort = 0
                    for ori in ortak_list:

                        if ort % 4 == 0:
                            ortak_ad_unvan.append(ori)
                        elif ort % 4 == 1:
                            ortak_pay.append(ori)
                        elif ort % 4 == 2:
                            ortak_payoran.append(ori)
                        elif ort % 4 == 3:
                            ortak_oyoran.append(ori)
                            ortak_pb.append('TL')
                        ort = ort + 1
                else:
                    ortak_ad_unvan.append(' ')
                    ortak_pay.append(' ')
                    ortak_payoran.append(' ')
                    ortak_pb.append(' ')
                    ortak_oyoran.append(' ')

            ortak_pd = pd.DataFrame({'Kod': kod,
                                     'Ortak Adı-Unvanı': ortak_ad_unvan,
                                     'Sermaye Payı': ortak_pay,
                                     'Sermaye Payı Oran': ortak_payoran,
                                     'Sermaye PB': ortak_pb,
                                     'Oy Oranı': ortak_oyoran,
                                     })
            return ortak_pd

        def DolayliOrtak(lis2, kod):
            """Sermayeye Dolaylı Yoldan Sahip Olan Gerçek ve Tüzel Kişiler kod-ad-pay TL-Pay oranı-oy oranı formatında pandasDataFrame listeler"""

            dolayli_ad_unvan = []
            dolayli_pay = []
            dolayli_payoran = []
            dolayli_pb = []

            if lis2.count('Son Durum İtibariyle Sermayeye Dolaylı Yoldan Sahip Olan Gerçek ve Tüzel Kişiler') > 0:
                dolayli_start = lis2.index(
                    'Son Durum İtibariyle Sermayeye Dolaylı Yoldan Sahip Olan Gerçek ve Tüzel Kişiler')
                dolayli_mini_list = lis2[dolayli_start:]
                if dolayli_mini_list[1] == 'Bilgi Mevcut Değil':
                    dolayli_ad_unvan.append(' ')
                    dolayli_pay.append(' ')
                    dolayli_payoran.append(' ')
                    dolayli_pb.append(' ')
                else:
                    if dolayli_mini_list.count('Fiili Dolaşımdaki Paylar') > 0:
                        dolayli_end = dolayli_mini_list.index('Fiili Dolaşımdaki Paylar')
                    else:
                        dolayli_end = dolayli_mini_list.index('Sermayeyi Temsil Eden Paylara İlişkin Bilgi')
                    dolayli_fin_list = dolayli_mini_list[5:dolayli_end]
                    do = 0
                    for dol in dolayli_fin_list:

                        if do % 4 == 0:
                            dolayli_ad_unvan.append(dol)
                        elif do % 4 == 1:
                            dolayli_pay.append(dol)
                        elif do % 4 == 2:
                            dolayli_pb.append(dol)
                        elif do % 4 == 3:
                            dolayli_payoran.append(dol)
                        do = do + 1

            dolayli_pd = pd.DataFrame({'Kod': kod,
                                       'Dolaylı Ortak Adı-Unvanı': dolayli_ad_unvan,
                                       'Sermaye Payı': dolayli_pay,
                                       'Sermaye PB': dolayli_pb,
                                       'Sermaye Payı Oran': dolayli_payoran,
                                       })
            return dolayli_pd

        def DolasimdakiPaylar(lis, kod):
            """Eğer firmanın pay senetleri borsada işlem görüyorsa kod-borsa kodu-dolaşındaki pay tutar-oran formatnda pandas dataFrame olarak listeler. dolaşımda payı yoksa da aynı formatta çıkarır."""
            dol_pay_kod = []
            dol_pay_tut = []
            dol_pay_oran = []
            if lis.count('Fiili Dolaşımdaki Pay Tutarı(TL)') > 0:
                dol_pay_start = lis.index('Borsa Kodu') + 3
                dol_pay_end = lis.index('Pay Grubu')
                dol_pay_list = lis[dol_pay_start:dol_pay_end]

                doli = 0
                for dol in dol_pay_list:
                    if doli % 3 == 0:
                        dol_pay_kod.append(dol)
                    elif doli % 3 == 1:
                        dol_pay_tut.append(dol)
                    elif doli % 3 == 2:
                        dol_pay_oran.append(dol)
                    doli = doli + 1
            else:
                dol_pay_kod.append(kod)
                dol_pay_tut.append(0)
                dol_pay_oran.append(0)

            dol_pay_pd = pd.DataFrame({'Kod': kod,
                                       'Borsa Kodu': dol_pay_kod,
                                       'Dolaşımdaki Pay Tutarı(TL)': dol_pay_tut,
                                       'Dolaşımdaki Pay Oranı': dol_pay_oran,
                                       })
            return dol_pay_pd

        def Istirakler(lis, kod):
            """Firma iştiraklerini/oratklıklarını kod-unvan-faaliyet konusu-sermaye-iştirak payı-para birimi-iştirak yüzdesi-ilişki niteliği formatında pandasDataFrame listeler"""

            istirak_unvan = []
            istirak_faal = []
            istirak_sermaye = []
            istirak_serm_pay = []
            istirak_parabir = []
            istirak_sirketpayi = []
            istirak_iliski = []

            if lis.count('Şirket ile Olan İlişkinin Niteliği') > 0:
                istirak_start = lis.index('Ticaret Unvanı') + 7
                istirak_end = len(lis)
                istirak_list = lis[istirak_start:istirak_end]

                isti = 0
                for ist in istirak_list:
                    if isti % 7 == 0:
                        istirak_unvan.append(ist)
                    elif isti % 7 == 1:
                        istirak_faal.append(ist)
                    elif isti % 7 == 2:
                        istirak_sermaye.append(ist)
                    elif isti % 7 == 3:
                        istirak_serm_pay.append(ist)
                    elif isti % 7 == 4:
                        istirak_parabir.append(ist)
                    elif isti % 7 == 5:
                        istirak_sirketpayi.append(ist)
                    elif isti % 7 == 6:
                        istirak_iliski.append(ist)
                    isti = isti + 1
            else:
                istirak_unvan.append('İştiraki Yok')
                istirak_faal.append('')
                istirak_sermaye.append('')
                istirak_serm_pay.append('')
                istirak_parabir.append('')
                istirak_sirketpayi.append('')
                istirak_iliski.append('')

            istirak_pd = pd.DataFrame({'Kod': kod,
                                       'İştirak Unvan': istirak_unvan,
                                       'İşt_Faaliyet Konusu': istirak_faal,
                                       'İşt_Sermayesi': istirak_sermaye,
                                       'Şirket_iştirak Payı': istirak_serm_pay,
                                       'Para Birimi': istirak_parabir,
                                       'Şirket_iştirak yüzde': istirak_sirketpayi,
                                       'İlişki Niteliği': istirak_iliski,
                                       })
            return istirak_pd

        yipd = yipd.append(Yat_ilis(ham_liste2, firma_kod), ignore_index=True)

        ykpd = ykpd.append(Yonetim_Kurulu(ham_liste, firma_kod), ignore_index=True)

        ysspd = ysspd.append(Yonetim_Soz_Sahibi(ham_liste, firma_kod), ignore_index=True)

        ortpd = ortpd.append(Ortaklar(ham_liste2, firma_kod), ignore_index=True)

        dolpd = dolpd.append(DolayliOrtak(ham_liste2, firma_kod), ignore_index=True)

        dppd = dppd.append(DolasimdakiPaylar(ham_liste, firma_kod), ignore_index=True)

        istpd = istpd.append(Istirakler(ham_liste, firma_kod), ignore_index=True)

        vit = vit + 1
        time.sleep(timeout)

    return yipd, ykpd, ysspd, ortpd, dolpd, dppd, istpd
"""
GenelDF = GenelBilgiler(firma_link_list, firma_kod_list,5)

with pd.ExcelWriter('Genel Bilgiler.xlsx') as writer:
    GenelDF[0].to_excel(writer, sheet_name='YatirimciIliskileri')
    GenelDF[1].to_excel(writer, sheet_name='YonetimKurulu')
    GenelDF[2].to_excel(writer, sheet_name='YonetimdeSozSahibi')
    GenelDF[3].to_excel(writer, sheet_name='Ortaklar')
    GenelDF[4].to_excel(writer, sheet_name='DolayliOrtaklar')
    GenelDF[5].to_excel(writer, sheet_name='DolasimdakiPaylar')
    GenelDF[6].to_excel(writer, sheet_name='Istirakler')
"""

