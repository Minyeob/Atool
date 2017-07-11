import re
import urllib.request
import xmltodict
from datetime import datetime

#인식한 엑셀파일을 바탕으로 API를 찾아 파이썬내에서 변수로 활용할 수 있게 하는 함수
class API_management():
#pnu뒤에 산 0-1확인
    def pnu_san(self, addr):
        if (addr[0].startswith('산') and addr[1].isnumeric() == True):
            return '2'  # 만약 산이 있으면
        else:
            return '1'  # 일반
    # 주소(리) 분리
    def pnu_split(self, addr):
        addr = addr.split()[0]
        addr = addr.split('-')
        if (len(addr) > 1):
            addr[0] = re.sub('[^0-9]', '', addr[0])
            addr[1] = re.sub('[^0-9]', '', addr[1])
            if (len(addr[0]) == 1):
                    addr[0] = '0' * 3 + addr[0]
            elif (len(addr[0]) == 2):
                 addr[0] = '0' * 2 + addr[0]
            elif (len(addr[0]) == 3):
                 addr[0] = '0' * 1 + addr[0]
            elif (len(addr[0]) == 4):
                addr[0] = addr[0]
            if (len(addr[1]) == 1):
                addr[1] = '0' * 3 + addr[1]
            elif (len(addr[1]) == 2):
                addr[1] = '0' * 2 + addr[1]
            elif (len(addr[1]) == 3):
                addr[1] = '0' * 1 + addr[1]
            elif (len(addr[1]) == 4):
                addr[1]=addr[1]
            return addr
        elif(len(addr)==1):
            addr = addr.append('0000')
            return addr
    # 주소(리) 분리/4자리 0 태그가 붙지 않는 경우
    def pnu_split_notag(self, addr):
                addr = addr.split()[0]
                addr = addr.split('-')
                if (len(addr) > 1):
                    for i in range(2):
                        addr[i] = re.sub('[^0-9]','',addr[i])
                return addr

            # fetch 토지임야정보조회서비스-지목추출

    #지목
    def fetch_LS_JM(self,pnu, code, addr_split, key):
        if(addr_split==None):
            return None
        else:
            url_LS = 'http://apis.data.go.kr/1611000/nsdi/eios/LadfrlService/ladfrlList.xml?pnu=' + str(pnu) + code + addr_split[0] + addr_split[1]+ '&ServiceKey=' + key
            data = urllib.request.urlopen(url_LS).read()
            check = xmltodict.parse(data)['fields']['totalCount']
            check = int(check)
            if (check == 0):
                return None
            elif(check > 0):
                data = xmltodict.parse(data)['fields']['ladfrlVOList']['lndcgrCodeNm']
        return data
    #층별 용도-대표용도 1개로 통일
    def fetch_LS_Main_PrupsCDNm(self,pnu, addr_split, key):
        pnu= str(pnu)
        pnu_1 = pnu[:5]
        pnu_2 = pnu[5:10]
        url_LS = 'http://apis.data.go.kr/1611000/BldRgstService/getBrRecapTitleInfo?sigunguCd='+pnu_1+'&bjdongCd='+pnu_2+'&bun='+addr_split[0]+'&ji='+addr_split[1]+'&ServiceKey=' + key
        data = urllib.request.urlopen(url_LS).read()
        check = xmltodict.parse(data)['response']['body']['totalCount']
        check = int(check)
        if(check==0):
            return None
        elif(check>=1):
                data = xmltodict.parse(data)['response']['body']['items']['item']['mainPurpsCdNm']
        return data
    #건물구조
    def fetch_LS_Main_strctCdNm(self, pnu, addr_split, key):
        pnu = str(pnu)
        pnu_1 = pnu[:5]
        pnu_2 = pnu[5:10]
        url_LS = 'http://apis.data.go.kr/1611000/BldRgstService/getBrTitleInfo?sigunguCd='+pnu_1+'&bjdongCd='+pnu_2+'&ServiceKey=' + key
        data = urllib.request.urlopen(url_LS).read()
        check = xmltodict.parse(data)['response']['body']['totalCount']
        check = int(check)
        if (check == 0):
            return None
        elif (check >= 1):
                data = xmltodict.parse(data)['response']['body']['items']['item'][0]['strctCdNm']
                if(data.find('구조')>0):
                    data = data[:len(data)-2]
                    return data
                else:
                    return data


        #사용승인일
    #사용승인일
    def fetch_LS_Main_crtnDay(self, pnu, addr_split, key):
        pnu = str(pnu)
        pnu_1 = pnu[:5]
        pnu_2 = pnu[5:10]
        url_LS = 'http://apis.data.go.kr/1611000/BldRgstService/getBrRecapTitleInfo?sigunguCd=' + pnu_1 +'&bjdongCd='+pnu_2+ '&ServiceKey=' + key
        data = urllib.request.urlopen(url_LS).read()
        check = xmltodict.parse(data)['response']['body']['totalCount']
        check = int(check)
        if (check == 0):
            return None
        elif (check >= 1):
            data = xmltodict.parse(data)['response']['body']['items']['item'][0]['crtnDay']
            return data
    # 전체토지면적
    def fetch_LS_Area(self, pnu, code, addr_split, key):
        if (addr_split == None):
            return None
        else:
            url_LS = 'http://apis.data.go.kr/1611000/nsdi/eios/LadfrlService/ladfrlList.xml?pnu=' + str(pnu) + code + addr_split[0] + addr_split[1] + '&ServiceKey=' + key
            data = urllib.request.urlopen(url_LS).read()
            check = xmltodict.parse(data)['fields']['totalCount']
            check = int(check)
            if (check == 0):
                return None
            elif (check > 0):
                data = xmltodict.parse(data)['fields']['ladfrlVOList']['lndpclAr']
        return data

    # 인근사례-용도
    def find_case_Purps(self, pnu, addr_split, key):
        pnu= str(pnu)
        url = 'http://apis.data.go.kr/1611000/BldRgstService/getBrTitleInfo?sigunguCd=' + pnu[:5]+ '&bjdongCd=' + pnu[5:10] + '&bun='+addr_split[0] + '&ji' + addr_split[1] +'&numOfRows=100'+ '&ServiceKey=' + key
        data = urllib.request.urlopen(url).read()
        for i in range(int(xmltodict.parse(data)['response']['body']['totalCount'])):
            if(xmltodict.parse(data)['response']['body']['items']['item'][i]['bjdongCd']==pnu[5:10] and xmltodict.parse(data)['response']['body']['items']['item'][i]['bun'] == addr_split[0]):
                output = xmltodict.parse(data)['response']['body']['items']['item'][i]['mainPurpsCdNm']
                return output
                break
            else:
                return '상가'
                break

    # 인근사례-구조
    def find_case_strCt(self, pnu, addr_split, key):
        pnu= str(pnu)
        if (addr_split == None):
            return None
        else:
            url = 'http://apis.data.go.kr/1611000/BldRgstService/getBrTitleInfo?sigunguCd=' + pnu[:5] + '&bjdongCd=' + pnu[5:10] + '&bun=' + addr_split[0] + '&ji' + addr_split[1] +'&numOfRows=100'+  '&ServiceKey=' + key
            data = urllib.request.urlopen(url).read()
            for i in range(int(xmltodict.parse(data)['response']['body']['totalCount'])):
                if (xmltodict.parse(data)['response']['body']['items']['item'][i]['bjdongCd'] == pnu[5:10] and
                            xmltodict.parse(data)['response']['body']['items']['item'][i]['bun'] == addr_split[0]):
                    output = xmltodict.parse(data)['response']['body']['items']['item'][i]['strctCdNm']
                    break
        return output

    # 인근사례=사용승인일
    def find_case_AprDay(self, pnu, addr_split, key):
        pnu= str(pnu)
        if (addr_split == None):
            return '2017-02-17'
        else:
            url = 'http://apis.data.go.kr/1611000/BldRgstService/getBrTitleInfo?sigunguCd=' + pnu[:5] + '&bjdongCd=' + pnu[5:10] + '&bun=' + addr_split[0] + '&ji' + addr_split[1] +'&numOfRows=100'+  '&ServiceKey=' + key
            data = urllib.request.urlopen(url).read()
            for i in range(int(xmltodict.parse(data)['response']['body']['totalCount'])):
                if (xmltodict.parse(data)['response']['body']['items']['item'][i]['bjdongCd'] == pnu[5:10] and
                            xmltodict.parse(data)['response']['body']['items']['item'][i]['bun'] == addr_split[0]):
                    output = xmltodict.parse(data)['response']['body']['items']['item'][i]['useAprDay']
                    break
            return output

    '''
    def find_case_trade(self,pnu,property_category,district,key):
        pnu = str(pnu)
        pnu = pnu[:5]
        date= datetime.today().strftime("%Y%m")
        if(property_category =="오피스텔(상가)"or property_category=="오피스텔(주거)"):
            url ='http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiTrade?LAWD_CD='+ pnu + '&DEAL_YMD=' +date+'&ServiceKey='+key
            data = urllib.request.urlopen(url).read
            data = xmltodict.parse(data)
            for i in range(int(totalcount = data['response']['body']['totalCount'])):
                if(data['response']['body']['items']['item'][i]['법정동'] == district):
                    s=[]
                    s.append(i)
            s=int(s)
            return data['response']['body']['items']['item'][s]

        elif(property_category =='연립'):
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcRHTrade?LAWD_CD='+ pnu + '&DEAL_YMD=' +date+'&ServiceKey='+key
            data = urllib.request.urlopen(url).read
            data = xmltodict.parse(data)
            for i in range(int(totalcount=data['response']['body']['totalCount'])):
                if (data['response']['body']['items']['item'][i]['법정동'] == district):
                    s = []
                    s.append(i)
            s = int(s)
            return data['response']['body']['items']['item'][s]

        else:
            return None'''




