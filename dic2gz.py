from selenium.common.exceptions import NoSuchElementException
import sys
import re
import csv
import datetime
import os
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import gzip
import shutil
import tempfile
import time

#   version 0.1.1

def scraping():
    # kb番号リスト
    kb_file = "kb.csv"
    # 前回分辞書ファイルall
    pre_af = "hotfix_all.csv"
    # 前回分辞書ファイルsmall
    pre_sf = "hotfix_small.csv"
    # osバージョン,ビルド番号,smallフラグリスト
    osb_f = "osbuild.csv"

    if not (os.path.exists(kb_file)):
        print('KB番号リスト', kb_file, 'が見つかりません。処理を終了します')
        return
        
    osb_dic = {}
    small_list = []
    if os.path.exists(osb_f):
        with open(osb_f, 'r', encoding='utf-8') as osf:
            oreader = csv.reader(osf)
            for osb_l in oreader:
                osv = osb_l[0]
                if osv[0] == '#':
                    continue
                osb_dic[osv] = osb_l[1]
                if len(osb_l) == 3:
                    if osb_l[2]=="1":
                        small_list.append(osv)


    with open(kb_file, 'r', encoding='utf-8') as kf:
        reader = csv.reader(kf)
        # 書込ファイルの指定
        now = datetime.datetime.now()
        tstamp = now.strftime('%Y%m%d%H%M%S')
        new_dir_path_all = (tstamp+"/all")
        new_dir_path_small = (tstamp+"/small")
        os.makedirs(new_dir_path_all, exist_ok=True)
        os.makedirs(new_dir_path_small, exist_ok=True)
        all_file = new_dir_path_all + '/hotfix_all.csv'
        small_file = new_dir_path_small + '/hotfix_small.csv'
        with open(all_file, 'a', encoding='utf-16', newline='') as wf:
            writer = csv.writer(wf, delimiter='\t',
                                quotechar='"', quoting=csv.QUOTE_ALL)
            if not (os.path.exists(pre_af)):
                print('前回分ファイル', pre_af, 'がないので辞書ファイル(all)を新規作成します')
            else:
                with open(pre_af, 'r', encoding='utf-16') as af:
                    for aline in af:
                        wf.write(aline)
            with open(small_file, 'a', encoding='utf-16', newline='') as wsf:
                swriter = csv.writer(wsf, delimiter='\t',
                                     quotechar='"', quoting=csv.QUOTE_ALL)
                if not (os.path.exists(pre_sf)):
                    print('前回分ファイル', pre_sf, 'がないので辞書ファイル(small)を新規作成します')
                else:
                    with open(pre_sf, 'r', encoding='utf-16') as presf:
                        for sline in presf:
                            wsf.write(sline)

                pre_kb = ""
                options = Options()
                # headlessオプションはつけない(つけると日本語取得不可)
                options.add_argument('--headless')

                # KB番号リストより順に値を取得
                for line in reader:
                    kb = line[0]
                    if kb[0] == "#":    # コメント行スキップ
                        continue

                    multiple_build = ""

                    # カタログページ
                    url = "https://www.catalog.update.microsoft.com/Search.aspx?q=KB" + kb
                    add_info18 = url  # カタログURL

                    url_times = 0
                    while True:
                        try:
                            driver1 = webdriver.Chrome()
                            driver1.get(url)
                            WebDriverWait(driver1, 30).until(EC.presence_of_all_elements_located)

                            if re.match('^.*問題が発生.*$', driver1.find_element_by_xpath('/html/body/div/form[2]/div[3]/table/tbody/tr[1]/td/div/div/div[1]/table/tbody/tr/td[2]/span').text):
                                print(kb, 'カタログページ再取得中')
                                url_times += 1
                                driver1.close()
                                if url_times == 3:
                                    print(kb, 'カタログページ取得エラー')
                                    break
                            else:
                                if url_times > 0:
                                    print(kb, 'カタログページ取得完了')
                                break

                        except:
                            print(kb, 'カタログページ取得エラー')
                            break

                    if url_times == 3:
                        break

                    # カタログページ数取得
                    page_item = driver1.find_element_by_id('ctl00_catalogBody_searchDuration').text.split('(1/')[1].split(' ')[0]
                    page_num = int(page_item)
                    for page in range(page_num):

                        pro_num = driver1.find_elements_by_class_name('flatBlueButtonDownload')    # 処理数獲得
                        for i in range(len(pro_num)):
                            # 辞書に書き込むためのリスト
                            print_info_list = []
                            # カタログページの項目リスト
                            item_list = []
                            # (必要KBの)カタログページの項目リスト(以下..._nd_kbは必要条件、重要の項目用)
                            item_list_nd_kb = []
                            arch = ""

                            # テーブル行取得
                            tbody_tr = driver1.find_element_by_css_selector(
                                'table#ctl00_catalogBody_updateMatches > tbody').find_elements_by_tag_name('tr')[i+1]

                            catalog_ele = tbody_tr.find_elements_by_tag_name(
                                'td')  # 列リスト取得

                            # カタログページの項目取得
                            for i_items in range(1, 7):
                                catalog_ele_text = catalog_ele[i_items].text.strip(
                                )

                                # 「製品」が'Windows 10,Windows 10 LTSB'の時'Windows 10'とする
                                if ",Windows 10 LTSB" in catalog_ele_text or ", Windows 10 LTSB" in catalog_ele_text:
                                    catalog_ele_2 = catalog_ele_text.split(',')[0]
                                    item_list.append(catalog_ele_2)
                                elif ", version" in catalog_ele_text:                  # 「製品」が'Windows 10, version 1903 and later'などのケース
                                    os_type = catalog_ele_text.split(',')[0]
                                    catalog_ele_2 = os_type
                                    item_list.append(catalog_ele_2)
                                else:
                                    # item_list [タイトル, 製品, 分類, 最終更新日時, バージョン, サイズ]
                                    item_list.append(catalog_ele_text)
                                
                            # ARM, Windows 10 GDR-DU 除くWindows 10
                            if "ARM" in item_list[0] or "Windows 10 GDR-DU" in item_list[1] or "Windows 10" not in item_list[1]:
                                continue
                            elif "差分" in item_list[0]:
                                print(kb, item_list[0], "は差分プログラムが対象外のためスキップします")
                                continue


                            add_info1 = "KB" + kb       # KB番号
                            add_info2 = item_list[1]    # 対象OS
                            add_info3 = "10.0.1.0"      # OSProduct

                            tnj_times = 0
                            while True:
                                if tnj_times > 0:
                                    driver1.switch_to.window(driver1.window_handles[0])
                                tbody_tr.find_element_by_tag_name('a').click()  # タイトルリンククリック
                                WebDriverWait(driver1, 30).until(EC.presence_of_all_elements_located)
                                driver1.switch_to.window(driver1.window_handles[1])

                                try:
                                    title_name_ja = driver1.find_element_by_id('ScopedViewHandler_titleText')
                                    if tnj_times > 0:
                                        print(kb, 'title_name_ja 取得完了')
                                    add_info10 = title_name_ja.text.strip()  # タイトル
                                    break

                                except NoSuchElementException:
                                    print(kb, 'title_name_ja 再取得中')
                                    tnj_times += 1
                                    if tnj_times == 3:
                                        print(kb, 'title_name_ja 取得エラー')
                                        break
                                    driver1.close()

                            if tnj_times == 3:
                                break

                            archtec = driver1.find_element_by_id('archDiv')
                            archtec_text = archtec.text.split(":")[1].strip()

                            if "x64" in add_info10:
                                arch = "AMD64"
                            elif "x86" in add_info10:
                                arch = "x86"
                            elif "," not in archtec_text:
                                arch = archtec_text
                            else:
                                print(kb, add_info10, archtec_text, 'アーキテクチャが特定できません')
                                break

                            add_info9 = arch                            # アーキテクチャ

                            # タイトルのスペース区切りのリストでOSのバージョンを探す
                            for i_ver in range(len(add_info10.split(' '))):
                                if re.match('^[0-9]{2}[H0-9][0-9]$', add_info10.split(' ')[i_ver]):    # ()なし
                                    add_info4 = add_info10.split(' ')[i_ver]                # OS Version
                                    break
                                elif re.match('^\([0-9]{2}[H0-9][0-9]\)$', add_info10.split(' ')[i_ver]):  # ()あり
                                    add_info4 = add_info10.split(' ')[i_ver].strip('('')')
                                    break
                                else:
                                    add_info4 = ""
                            
                            if add_info4 not in osb_dic:
                                print(kb,add_info10, "は対象外のOSバージョンのためスキップします")
                                driver1.close()
                                driver1.switch_to.window(driver1.window_handles[0])
                                continue

                            # 最終更新日時
                            add_info14 = item_list[3]

                            classfication = driver1.find_element_by_id('classificationDiv')
                            add_info24 = classfication.text.split(':')[1].strip()           # 分類

                            # バージョン "(N/A)"
                            add_info15 = item_list[4]

                            size = item_list[5].strip('MB ')
                            if "K" in size:
                                size = size.strip('KB')
                                add_info16 = int(size)/1000
                            else:
                                add_info16 = size                   # サイズ

                            sec_num = driver1.find_element_by_id('securityBullitenDiv')
                            add_info19 = sec_num.text.split(':')[1].strip()                 # セキュリティ番号

                            sec_sev = driver1.find_element_by_id('msrcSeverityDiv')
                            add_info20 = sec_sev.text.split(':')[1].strip()

                            add_info21 = kb                                                 # サポート技術番号

                            det_url = driver1.find_element_by_css_selector('div#moreInfoDiv a')
                            add_info22 = det_url.text                                       # 詳細URL

                            sup_url = driver1.find_element_by_css_selector('div#suportUrlDiv a')
                            add_info23 = sup_url.text                                      # サポートURL

                            driver1.close()                                                 # 詳細画面クローズ

                            driver1.switch_to.window(driver1.window_handles[0])
                            driver1.find_elements_by_class_name('flatBlueButtonDownload')[i].click()   # ダウンロードボタンクリック
                            WebDriverWait(driver1, 30).until(EC.presence_of_all_elements_located)
                            driver1.switch_to.window(driver1.window_handles[1])
                            # 次の処理は任意#################################################################################33
                            time.sleep(1)

                            furl_times = 0
                            while True:
                                try:
                                    element_furl = driver1.find_element_by_tag_name('a').get_attribute('href')
                                    if furl_times > 0:
                                        print(kb, add_info10, '取得成功')
                                    break

                                except NoSuchElementException:
                                    print(kb, add_info10, 'のelement_furlの要素が見つかりません。再取得を試みます')
                                    driver1.refresh()
                                    WebDriverWait(driver1, 30).until(EC.presence_of_all_elements_located)
                                    furl_times += 1
                                    if furl_times == 3:
                                        print('3回トライしましたがダウンロードURLを取得できませんでした')
                                        break

                            if furl_times == 3:
                                break
                            # ダウンロードURL
                            add_info11 = element_furl

                            ftl_times = 0
                            while True:
                                try:
                                    element_ftitle = driver1.find_element_by_class_name('textTopTitlePadding.textBold.textSubHeadingColor')
                                    if ftl_times > 0:
                                        print(kb, 'element_ftitle取得完了')
                                    add_info12 = element_ftitle.text  # ダウンロードタイトル
                                    break

                                except NoSuchElementException:
                                    print(kb, 'element_ftitle再取得中')
                                    ftl_times += 1
                                    if ftl_times == 3:
                                        print(kb, 'element_ftitle取得エラー')
                                        break

                            if ftl_times == 3:
                                break

                            fnm_times = 0
                            while True:
                                try:
                                    element_fname = driver1.find_element_by_tag_name('a')
                                    if fnm_times > 0:
                                        print(kb, 'element_fname取得完了')
                                    break
                                except NoSuchElementException:
                                    print(kb, 'element_fname再取得中')
                                    fnm_times += 1
                                    if fnm_times == 3:
                                        print(kb, 'element_fname取得エラー')
                                        break
                            if fnm_times == 3:
                                break
                            # ダウンロードファイル名
                            add_info13 = element_fname.text

                            # ダウンロードページクローズ
                            driver1.close()
                            driver1.switch_to.window(driver1.window_handles[0])

                            # 前のKB番号と異なる or ビルド番号を2つもつKB番号の場合、以下を処理
                            if (pre_kb != kb) or (multiple_build == 1 and add_info4 != pre_add_info4):
                                burl_times = 0
                                while True:
                                    try:
                                        driver2 = webdriver.Chrome(options=options)                     # サポートページ
                                        build_url = "https://support.microsoft.com/ja-jp/help/" + \
                                            kb + "/windows-10-update-kb" + kb
                                        driver2.get(build_url)
                                        WebDriverWait(driver2, 30).until(EC.presence_of_all_elements_located)
                                        if burl_times > 0:
                                            print(kb, 'build_url取得完了')
                                        break

                                    except NoSuchElementException:
                                        print(kb, 'build_url再取得中')
                                        burl_times += 1
                                        if burl_times == 3:
                                            print(kb, 'build_url取得エラー')
                                            break
                                        driver2.close()

                                if burl_times == 3:
                                    break

                                bld_times = 0
                                while True:
                                    try:
                                        element_build = driver2.find_element_by_tag_name('h1')
                                        if bld_times > 0:
                                            print(kb, 'element_build取得完了')
                                        break
                                    except NoSuchElementException:
                                        print(kb, 'element_build再取得中')
                                        bld_times += 1
                                        if bld_times == 3:
                                            print(kb, 'element_build取得エラー')
                                            break
                                        driver2.refresh()
                                        WebDriverWait(driver2, 30).until(EC.presence_of_all_elements_located)

                                if bld_times == 3:
                                    break

                                element_build_text = element_build.text                             # ページタイトル

                                if "." in element_build_text:
                                    # ビルド番号が２つあるケース
                                    if ("および" in element_build_text) or ("and" in element_build_text):
                                        xps_times = 0
                                        while True:
                                            try:
                                                build = driver2.find_element_by_css_selector('#supArticleContent > div > div > div > div.supARG-column-2-3 > p:nth-child(2) > b')  # バージョン行
                                                if xps_times > 0:
                                                    print(kb, 'build取得完了')
                                                break

                                            except NoSuchElementException:
                                                print(kb, 'build再取得中')
                                                xps_times += 1
                                                if xps_times == 3:
                                                    print(kb, 'build取得エラー')
                                                    break
                                                driver2.refresh()
                                                WebDriverWait(driver2, 30).until(EC.presence_of_all_elements_located)

                                        if xps_times == 3:
                                            break
                                        build_text = build.text
                                        if " - OS ビルド" in build_text:
                                            sep = add_info4 + " - OS ビルド"
                                        elif "-OS ビルド" in build_text:
                                            sep = add_info4 + "-OS ビルド"
                                        elif " - OS Build" in build_text:
                                            sep = add_info4 + " - OS Build"
                                        elif "-OS Build" in build_text:
                                            sep = add_info4 + "-OS Build"
                                        else:
                                            sep = "OS Build"
                                        build = build_text.split(sep)[1].strip().split(" ")[0]
                                        build2_item = ""
                                        if "、" in build:
                                            build2_item = build.split("、")[0].strip()
                                        else:
                                            build2_item = build
                                        add_info5 = build.split('.')[0]                         # build1
                                        add_info6 = build2_item.split('.')[1]                   # build2
                                        if add_info5 == add_info6:
                                            add_info6 = build2_item.split('.')[2]        # "KB4580386ビルド番号誤記対応：18363.18363.1171"
                                        multiple_build = 1                       # 複数ビルドフラグ
                                        pre_add_info4 = add_info4           # OSVersionセット
                                    elif "(OS ビルド " in element_build_text:               # ()あり(日本語)
                                        build = element_build.text.split(
                                            "(OS ビルド ")[1]
                                        add_info5 = build.split('.')[0]
                                        add_info6 = build.split(
                                            '.')[1].split(')')[0]
                                    elif "（OS ビルド " in element_build_text:               # 全角()あり(日本語)
                                        build = element_build.text.split(
                                            "（OS ビルド ")[1]
                                        add_info5 = build.split('.')[0]
                                        add_info6 = build.split(
                                            '.')[1].split('）')[0]
                                    elif "(OS Build " in element_build_text:                # ()あり(英語)
                                        build = element_build.text.split(
                                            "(OS Build ")[1]
                                        add_info5 = build.split('.')[0]
                                        add_info6 = build.split(
                                            '.')[1].split(')')[0]
                                    elif "ビルド " in element_build_text:
                                        build = element_build.text.split("ビルド ")[
                                            1]
                                        add_info5 = build.split('.')[0]
                                        add_info6 = build.split('.')[1]
                                    else:
                                        add_info5 = ""
                                        add_info6 = ""
                                        print(kb, 'os build が取得できません')
                                        break
                                elif "Security update" or "セキュリティ更新" in element_build_text:
                                    if add_info4 in osb_dic:
                                        add_info5 = osb_dic[add_info4]
                                        add_info6 = ""
                                        multiple_build = 1                       # 複数ビルドフラグ
                                        pre_add_info4 = add_info4           # OSVersionセット
                                    else:
                                        print(kb,add_info10, 'os build が取得できません')
                                        add_info5 = ""
                                        add_info6 = ""
                                        break
                                else:
                                    print(kb,add_info10, 'os build が取得できません')
                                    add_info5 = ""
                                    add_info6 = ""
                                    break

                                try:
                                    element_condition_h2 = driver2.find_elements_by_tag_name('h2')
                                except NoSuchElementException:
                                    print(kb, 'element_condition_h2の要素が見つかりません。idが変更された可能性があります')
                                    break
                                try:
                                    element_condition_p = driver2.find_elements_by_tag_name('p')
                                except NoSuchElementException:
                                    print(kb, 'element_condition_pの要素が見つかりません。idが変更された可能性があります')
                                    break

                                need_kb_num = 0
                                start_state = 0
                                switch1 = "off"
                                switch2 = "off"

                                add_info7 = ""  # 必要KB
                                add_info8 = ""  # 必要条件/重要
                                for q in range(len(element_condition_h2)):
                                    # 英語対応要になった時のためのメモ：この更新プログラムの入手方法="How to get this update"(ex.KB4494441) en-usサイト参照
                                    if re.match('この更新プログラムの入手方法', element_condition_h2[q].text):
                                        start_state = element_condition_h2[q].location['y']
                                        switch1 = "on"
                                        break
                                if switch1 == "on":
                                    for r in range(len(element_condition_p)):
                                        # 英語対応要になった時のためのメモ：必要条件=Prerequisite(ex.KB4494441), 重要=Important(ex.KB4093112), en-usサイト参照
                                        if re.match('必要条件', element_condition_p[r].text) or re.match('重要', element_condition_p[r].text):
                                            if element_condition_p[r].location['y'] > start_state:
                                                add_info8 = element_condition_p[r].text.strip(
                                                )
                                                # ()あり
                                                if re.match('^.*\(SSU\) \(KB[0-9]{7}\).*$', add_info8):
                                                    kb_state = element_condition_p[r].text.find(
                                                        '(SSU)')
                                                    need_kb_num = element_condition_p[r].text[kb_state +
                                                                                                9: kb_state + 16]
                                                    add_info7 = "KB" + \
                                                        str(need_kb_num)
                                                    switch2 = "on"
                                                    break
                                                # ()なし
                                                elif re.match('^.*\(SSU\) KB[0-9]{7}.*$', add_info8):
                                                    kb_state = element_condition_p[r].text.find(
                                                        '(SSU)')
                                                    need_kb_num = element_condition_p[r].text[kb_state +
                                                                                                8: kb_state + 15]
                                                    add_info7 = "KB" + \
                                                        str(need_kb_num)
                                                    switch2 = "on"
                                                    break
                                                else:
                                                    print(kb, '必要KB')
                                driver2.close()

                                add_info17 = ""  # 必要KBリリース日
                                if switch2 == "on":
                                    driver3 = webdriver.Chrome(options=options)
                                    need_url = "https://www.catalog.update.microsoft.com/Search.aspx?q=KB" + need_kb_num
                                    driver3.get(need_url)
                                    WebDriverWait(driver3, 30).until(EC.presence_of_all_elements_located)
                                    catalog_ele_nd_kb = driver3.find_element_by_xpath(
                                        '/html/body/div/form[2]/div[3]/table/tbody/tr[1]/td/div/div/div[2]/table/tbody').find_elements_by_tag_name(
                                        'tr')[1].find_elements_by_tag_name('td')
                                    # 必要条件KBカタログの1行目の最終更新日時を取得
                                    for i_items_nd_kb in range(1, 7):
                                        if ',' in catalog_ele_nd_kb[i_items_nd_kb].text:
                                            catalog_ele_nd_kb_1 = catalog_ele_nd_kb[i_items_nd_kb].text.strip().split(',')[
                                                0]
                                            item_list_nd_kb.append(
                                                catalog_ele_nd_kb_1)
                                        else:
                                            item_list_nd_kb.append(
                                                catalog_ele_nd_kb[i_items_nd_kb].text.strip())

                                    f_date_nd_kb = item_list_nd_kb[3].split(
                                        '/')
                                    add_info17 = f_date_nd_kb[2] + "/" + \
                                        f_date_nd_kb[0] + "/" + f_date_nd_kb[1]

                                    driver3.close()

                            for info_num in range(1, 25):   # 取得情報書き込み
                                add_info = eval("add_info" + str(info_num))
                                print_info_list.append(add_info)

                            writer.writerow(print_info_list)

                            for small_os in small_list:
                                if add_info4 == small_os:
                                    swriter.writerow(print_info_list)
                                    break

                            pre_kb = kb

                        # 次のページがあれば「次へ」をクリック
                        if (page_num > 1) and (page < page_num - 1):
                            driver1.find_element_by_link_text("次へ").click()
                            WebDriverWait(driver1, 30).until(EC.presence_of_all_elements_located)

                    driver1.quit()

    header = '"1"\t"'+tstamp+'"\n'  # ヘッダー(1,timestamp)
    with open(all_file, 'r', encoding="utf-16") as gzaf:  # 辞書(all)読み込み
        # ヘッダー追加用に一時ファイルを作成
        with tempfile.NamedTemporaryFile(mode="w+", newline="", encoding="utf-16", delete=False) as afp:
            afp_name = afp.name
            afp.write(header)  # ヘッダーを1行目に書き込み
            for line in gzaf:
                afp.write(line)  # 辞書書き込み

        with open(afp_name, "rb") as f_in:  # 一時ファイルを読み込みモードで開く
            with gzip.open(new_dir_path_all + '/HotFix.url.gz', 'wb') as f_out:  # gzファイル書き込み
                shutil.copyfileobj(f_in, f_out)

        os.unlink(afp_name)  # 一時ファイルの削除

    with open(small_file, 'r', encoding="utf-16") as gzsf:  # 辞書(small)読み込み
        # ヘッダー追加用に一時ファイルを作成
        with tempfile.NamedTemporaryFile(mode="w+", newline="", encoding="utf-16", delete=False) as sfp:
            sfp_name = sfp.name
            sfp.write(header)  # ヘッダーを1行目に書き込み
            for line in gzsf:
                sfp.write(line)  # 辞書書き込み

        with open(sfp_name, "rb") as f_in:  # 一時ファイルを読み込みモードで開く
            with gzip.open(new_dir_path_small + '/HotFix.url.gz', 'wb') as f_out:  # gzファイル書き込み
                shutil.copyfileobj(f_in, f_out)

        os.unlink(sfp_name)  # 一時ファイルの削除


if __name__ == '__main__':
    start = time.time()
    scraping()
    goal = time.time()
    print("time(s):", round((goal-start), 2))
