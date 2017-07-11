from django.http import HttpResponseRedirect, HttpResponse
from django.shortcuts import render,redirect,get_object_or_404
from django.views.generic import *
#  파일을 import 할 때 from 에서 .을 이용하면 파일경로를 전부 칠 필요없이 현재 파일이 속한 파일의 다른 파일들을 가져올 수 있다
from .forms import UploadFileForm, ExcelForm
from .models import Document
from django.core.urlresolvers import reverse_lazy
from django.template import RequestContext
from .function import *
from .API_function import *
import xlrd
from django import shortcuts
from django.core.files import File
import os
import urllib.request
import xmltodict
import django.utils.http
import urllib

# -*- coding: utf-8 -*-

#get으로 해당 페이지에 접속하면 파일을 업로드할 수 있는 폼을 제공, 파일을 업로드해서 업로드버튼을 누르면 해당 파일에 있는 데이터로 각 사건에 대해 보기를 제공
def upload_file(request):
    if request.method == 'POST':
        form=UploadFileForm(request.POST, request.FILES)

        #자신이 만든 폼의 필드는 기본값으로 required=true 로 되어 있으므로 모든 필드가 입력되지 않으면 유효하지 않다
        if form.is_valid():
            new_document=Document(file=request.FILES['file'])
            new_document.title=new_document.file.name
            workbook =excel_handling().make_file(new_document.file)
            normal_datas=excel_handling().get_normal(workbook)
            normal_codes=excel_handling().get_normal_code(workbook)
            normal_zip=zip(normal_datas, normal_codes)
            file=Document.objects.filter(title=new_document.title)
            animal = ['cat', 'dog', 'mause']
            how_many = ['one', 'two', 'three']
            data = zip(animal, how_many)

            #해당파일이 이미 존재하면 저장하지 않고 해당파일이 없다면 해당 파일의 데이터 모델을 저장한다
            if(len(file)==0):
                new_document.save()

            return render(request,'ibk/templatemo_497_upper/templatemo_497_upper/code_selection.html', {'normals':normal_zip, 'file':new_document, 'codes':data})

    else:
        form=UploadFileForm()

    return render(request,'ibk/templatemo_497_upper/templatemo_497_upper/form.html',{'form':form})

#탁감 보고서를 작성하기 위해 엑셀데이터에서 필요한 데이터를 추출해서 html 페이지로 데이터를 보내준다
def show_normal_report(request, code):
    name=request.GET['title']
    print(name)
    file=get_object_or_404(Document, title=name)
    loc=0;
    workbook = excel_handling().make_file(file.file)
    #모든 데이터로 데이터 집합을 구해 보통 같은 행의 데이터는 하나의 대상에 대한 곧통의 데이터를 가르키므로 해당 행의 위치를 구해 나머지도 구한다
    codes = excel_handling().get_all_code(workbook)
    while(loc<len(codes)):
        if(codes[loc]==code):
            break
        else:
            loc=loc+1
    borrow_name=excel_handling().get_render_name(workbook, loc)    #Borrow Name
    program_title=excel_handling().get_program_title(workbook)  #Program
    property_control_no=excel_handling().get_property_control_no(workbook,loc) #Property Control No
    court=excel_handling().get_court(workbook,loc)  #관할법원
    case=excel_handling().get_case_number(workbook, loc)    #사건번호
    borrower_num=excel_handling().get_render_index(workbook, loc)   #차주일련번호
    opb=excel_handling().get_opb(workbook, borrower_num)    #OPB
    interest=excel_handling().get_accured_interest(workbook, borrower_num)  #연체이자
    credit_amount=opb+interest  #총 채권액
    setup_price=excel_handling().get_cpma(workbook, loc) #설정액
    address=excel_handling().get_address(workbook, loc, code)    #Address
    category=excel_handling().get_property_category(workbook, loc)  #Property category
    label=['가','나','다','라','마','바','사','아','자','차','카','타','파','하']
    ho=excel_handling().get_ho(workbook, code)  #건물의 호들
    liensize_improvement=excel_handling().get_liensize_improvement(workbook, code)  #전유면적 - 일단 엑셀에서 건물면적을 꺼내서 구함
    liensize_improvement_py=[]
    for index in liensize_improvement:
        index=round(float(index)*121/400,2)
        liensize_improvement_py.append(index)

    landsize=excel_handling().get_landsize(workbook, code)  #대지권 면적
    landsize_py=[]
    for index in landsize:
        index=round(float(index)*121/400,2)
        landsize_py.append(index)
    sum_ho=len(ho)
    sum_liensize_improvement=0
    sum_liensize_improvement_py= 0
    sum_landsize=0
    sum_landsize_py= 0
    for a in liensize_improvement:
        sum_liensize_improvement=sum_liensize_improvement+round(float(a),2)
    for a in liensize_improvement_py:
        sum_liensize_improvement_py=sum_liensize_improvement_py+round(float(a),2)
    for a in landsize:
        sum_landsize=sum_landsize+round(float(a),2)
    for a in landsize_py:
        sum_landsize_py=sum_landsize_py+round(float(a),2)
    sum_liensize_improvement=round(sum_liensize_improvement,2)
    sum_liensize_improvement_py=round(sum_liensize_improvement_py,2)
    sum_landsize=round(sum_landsize,2)
    sum_landsize_py=round(sum_landsize_py,2)


    building = zip(label, ho, liensize_improvement, liensize_improvement_py, landsize, landsize_py)
    utensil = excel_handling().get_utensil(workbook, loc)  # 기계기구의 숫자
    address_code=excel_handling().get_address_code(workbook,loc)

    #시,구,동
    entire_address=address.split()
    si=entire_address[0]
    gu=entire_address[1]
    dong=entire_address[2]

    #url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcSHRent?LAWD_CD=11110&DEAL_YMD=201702&serviceKey=KqP4dQTZbN2QbXZlZUK0gYsfRfqiACwnmgqPf3N2yPqj%2F7Ura0eDpY1CKVPmzzQRqGS3myGv3Oauhw7YmfPDLg%3D%3D'
    #data = urllib.request.urlopen(url).read()
    #sample=xmltodict.parse(data)['response']['body']['items']['item'][1]['계약면적']

    #download_url=excel_write().save_file(program_title, opb, property_control_no, interest, setup_price)
    '''APT'''
    address_later = excel_handling().get_adddistrict(workbook, loc, code)  # 법정'동'
    address_last = excel_handling().get_addtdistrict(workbook, loc, code)  # 법정동 마지막 주소
    pnu_san = API_management().pnu_san(address_last)  # pnu-일반: 1 산: 2
    pnu_split = API_management().pnu_split(address_last)  # 법정동 마지막 주소로부터 지번과 부번-pnu 19자리에 사용
    pnu_split_notag = API_management().pnu_split_notag(address_last)  # 법정동 주소로부터 지번과 부번 추출

    # 지목코드 추출
    urlfor_lnd_info = API_management().fetch_LS_JM(address_code, pnu_san, pnu_split, key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # 층별용도 추출
    urlfor_lnd_mainPurps = API_management().fetch_LS_Main_PrupsCDNm(address_code, pnu_split, key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # 건물구조
    urlfor_lnd_crt = API_management().fetch_LS_Main_strctCdNm(address_code, pnu_split_notag,
                                                              key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # 사용승인일
    urlfor_lnd_date = API_management().fetch_LS_Main_crtnDay(address_code, pnu_split_notag,
                                                             key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # 전체토지면적
    urlfor_lnd_area = API_management().fetch_LS_Area(address_code, pnu_san, pnu_split,
                                                     key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # 전체토지면적_py
    urlfor_lnd_area_py = round(float(urlfor_lnd_area)*0.3025,2)
    # 인근사례-용도
    urlfor_near_Purps = API_management().find_case_Purps(address_code, pnu_split,
                                                         key='KqP4dQTZbN2QbXZlZUK0gYsfRfqiACwnmgqPf3N2yPqj%2F7Ura0eDpY1CKVPmzzQRqGS3myGv3Oauhw7YmfPDLg%3D%3D')
    # 인근사례-구조
    urlfor_near_strCt = API_management().find_case_strCt(address_code, pnu_split,
                                                         key='KqP4dQTZbN2QbXZlZUK0gYsfRfqiACwnmgqPf3N2yPqj%2F7Ura0eDpY1CKVPmzzQRqGS3myGv3Oauhw7YmfPDLg%3D%3D')
    # 인근사례-사용승인일
    urlfor_near_AprDay = API_management().find_case_AprDay(address_code, pnu_split,
                                                           key='KqP4dQTZbN2QbXZlZUK0gYsfRfqiACwnmgqPf3N2yPqj%2F7Ura0eDpY1CKVPmzzQRqGS3myGv3Oauhw7YmfPDLg%3D%3D')

    #  near_dong = API_management().find_case_trade(address_code, category, address_later,
                                             #    key='bbxaimTqNXp3yIqt1L2B0s7bGtNJ%2F2BBx7qeN15NzQJOsO%2BXiQXpk6oPP9eF%2B%2FpgHIbl4LrafbsyPhoaN889gg%3D%3D')
    # near_lnd_crt = API_management().fetch_LS_Main_strctCdNm(near_dong[''])
    # download_url=excel_write().save_file(program_title, opb, property_control_no, interest, setup_price)

    return render(request, 'ibk/report.html',
                  {'code':code,'borrow_name':borrow_name, 'program':program_title, 'property_control_no':property_control_no, 'court':court, 'case':case, 'opb':opb,
                   'interest':interest, 'credit_amout':credit_amount, 'setup_price':setup_price, 'address':address, 'category':category, 'building':building,
                   'utensil':utensil, 'address_code':address_code, 'sum_ho':sum_ho, 'sum_liensize_improvement':sum_liensize_improvement,
                   'urlfor_lnd_info': urlfor_lnd_info, 'urlfor_lnd_crt': urlfor_lnd_crt,
                   'urlfor_lnd_mainPurps': urlfor_lnd_mainPurps, 'urlfor_lnd_date': urlfor_lnd_date, 'urlfor_lnd_area' : urlfor_lnd_area, 'urlfor_lnd_area_py': urlfor_lnd_area_py,
                   'sum_liensize_improvement_py':sum_liensize_improvement_py, 'sum_landsize':sum_landsize, 'sum_landsize_py':sum_landsize_py,
		   'urlfor_near_Purps': urlfor_near_Purps, 'urlfor_near_strCt': urlfor_near_strCt, 'urlfor_near_AprDay':urlfor_near_AprDay,
                   'si':si, 'gu':gu, 'dong':dong})

def download(request):
    if request.method == 'POST':
        form = ExcelForm(request.POST)
        #form은 valid 검사를 하고 나서야 cleaned_data를 가진다
        if form.is_valid():
            #dictionary 형태로 선언하기 위해서는 중괄호 {} 로 선언해야 한다
            user_input={}
            #결과요약
            user_input['program'] = form.cleaned_data['program']
            user_input['opb'] = form.cleaned_data['opb']
            user_input['interest'] = form.cleaned_data['interest']
            user_input['property_control_no'] = form.cleaned_data['property_control_no']
            user_input['setup_price'] = form.cleaned_data['setup_price']
            user_input['user'] = form.cleaned_data['user']
            user_input['user_phone'] = form.cleaned_data['user_phone']  # 담당자 연락처
            user_input['credit_amount'] = form.cleaned_data['credit_amount'] # 총 채권액
            user_input['borrow_name'] = form.cleaned_data['borrow_name']  # borrow name
            user_input['law_price'] = form.cleaned_data['law_price']  # 법사가
            user_input['market_predict'] = form.cleaned_data['market_predict']  # 시장전망
            user_input['market_price'] = form.cleaned_data['market_price']  # 시장가
            user_input['law_price_comp1'] = form.cleaned_data['law_price_comp1']  # 법사가 대비 1
            user_input['market_price_comp1'] = form.cleaned_data['market_price_comp1']    # 시장가 대비 1
            user_input['opb_comp1'] = form.cleaned_data['opb_comp1']  # opb 대비 1
            user_input['court'] = form.cleaned_data['court']  # 관할법원
            user_input['bid'] = form.cleaned_data['bid']  # 낙찰가
            user_input['law_price_comp2'] = form.cleaned_data['law_price_comp2']  # 법사가 대비 2
            user_input['market_price_comp2'] = form.cleaned_data['market_price_comp2']  # 시장가 대비 2
            user_input['opb_comp2'] = form.cleaned_data['opb_comp2']  # opb 대비 2
            user_input['case'] = form.cleaned_data['case']  # 사건번호
            user_input['avg_bid'] = form.cleaned_data['avg_bid']  # 평균낙찰가
            user_input['law_price_comp3'] = form.cleaned_data['law_price_comp3']  # 법사가 대비 3
            user_input['market_price_comp3'] = form.cleaned_data['market_price_comp3']  # 시장가 대비 3
            user_input['opb_comp3'] = form.cleaned_data['opb_comp3']  # opb 대비 3
            user_input['submission_date'] = form.cleaned_data['submission_date']  # 법원제출일
            user_input['next_date'] = form.cleaned_data['next_date']  # 차기기일
            user_input['fail_count'] = form.cleaned_data['fail_count']  # 유찰회수

            #본건현황
            user_input['address'] = form.cleaned_data['address']  # Address
            user_input['property_category'] = form.cleaned_data['property_category']  # Property category
            user_input['usage'] = form.cleaned_data['usage']  # 용도지역
            user_input['urlfor_lnd_info'] = form.cleaned_data['urlfor_lnd_info']  # 지목
            user_input['state'] = form.cleaned_data['state']  # 이용상황
            user_input['land_price_m'] = form.cleaned_data['land_price_m']  # 개별공시지가 m
            user_input['land_price_py'] = form.cleaned_data['land_price_py']  # 개별공시지가 py
            user_input['urlfor_lnd_area'] = form.cleaned_data['urlfor_lnd_area']  # 전체토지면적 m
            user_input['land_size_py'] = form.cleaned_data['land_size_py']  # 전체토지면적 py
            user_input['security_size_m'] = form.cleaned_data['security_size_m']  # 담보면적 m
            user_input['security_size_py'] = form.cleaned_data['security_size_py']  # 담보면적 py
            user_input['urlfor_lnd_crt'] = form.cleaned_data['urlfor_lnd_crt']  # 건물 구조
            user_input['urlfor_lnd_date'] = form.cleaned_data['urlfor_lnd_date']  # 사용승인일
            user_input['urlfor_lnd_mainPurps'] = form.cleaned_data['urlfor_lnd_mainPurps']  # 층별 용도
            user_input['exclusive_rate'] = form.cleaned_data['exclusive_rate']  # 전용율
            user_input['exclusive_area_m'] = form.cleaned_data['exclusive_area_m']  # 전유면적 m
            user_input['exclusive_area_py'] = form.cleaned_data['exclusive_area_py']  # 전유면적 py
            user_input['contract_area_m'] = form.cleaned_data['contract_area_m']  # 계약면적 m
            user_input['contract_area_py'] = form.cleaned_data['contract_area_py']  # 계약면적 py

            #건물
            user_input['building_label']=request.POST.getlist('building_label[]')
            user_input['building_ho'] = request.POST.getlist('building_ho[]')
            user_input['building_exclusive_m']=request.POST.getlist('building_exclusive_m[]')
            user_input['building_exclusive_py'] = request.POST.getlist('building_exclusive_py[]')
            user_input['building_contract_m'] = request.POST.getlist('building_contract_m[]')
            user_input['building_contract_py'] = request.POST.getlist('building_contract_py[]')
            user_input['building_right_m'] = request.POST.getlist('building_right_m[]')
            user_input['building_right_py'] = request.POST.getlist('building_right_py[]')
            user_input['building_ratio'] = request.POST.getlist('building_ratio[]')
            user_input['building_auction_price'] = request.POST.getlist('building_auction_price[]')
            user_input['building_auction_exclusive'] = request.POST.getlist('building_auction_exclusive[]')
            user_input['building_auction_contract'] = request.POST.getlist('building_auction_contract[]')
            user_input['building_auction_ratio'] = request.POST.getlist('building_auction_ratio[]')
            user_input['building_market_price'] = request.POST.getlist('building_market_price[]')
            user_input['building_market_exclusive'] = request.POST.getlist('building_market_exclusive[]')
            user_input['building_market_contract'] = request.POST.getlist('building_market_contract[]')
            user_input['building_market_ma'] = request.POST.getlist('building_market_ma[]')
            user_input['building_estimated_price'] = request.POST.getlist('building_estimated_price[]')
            user_input['building_estimated_exclusive'] = request.POST.getlist('building_estimated_exclusive[]')
            user_input['building_estimated_contract'] = request.POST.getlist('building_estimated_contract[]')
            user_input['building_estimated_ea'] = request.POST.getlist('building_estimated_ea[]')
            user_input['building_estimated_em'] = request.POST.getlist('building_estimated_em[]')

            user_input['summary_ho'] = form.cleaned_data['summary_ho']
            user_input['summary_exclusive_m'] = form.cleaned_data['summary_exclusive_m']
            user_input['summary_exclusive_py'] = form.cleaned_data['summary_exclusive_py']
            user_input['summary_contract_m'] = form.cleaned_data['summary_contract_m']
            user_input['summary_contract_py'] = form.cleaned_data['summary_contract_py']
            user_input['summary_right_m'] = form.cleaned_data['summary_right_m']
            user_input['summary_right_py'] = form.cleaned_data['summary_right_py']
            user_input['summary_ratio'] = form.cleaned_data['summary_ratio']
            user_input['summary_auction_price'] = form.cleaned_data['summary_auction_price']
            user_input['summary_auction_exclusive'] = form.cleaned_data['summary_auction_exclusive']
            user_input['summary_auction_contract'] = form.cleaned_data['summary_auction_contract']
            user_input['summary_auction_ratio'] = form.cleaned_data['summary_auction_ratio']
            user_input['summary_market_price'] = form.cleaned_data['summary_market_price']
            user_input['summary_market_exclusive'] = form.cleaned_data['summary_market_exclusive']
            user_input['summary_market_contract'] = form.cleaned_data['summary_market_contract']
            user_input['summary_market_ma'] = form.cleaned_data['summary_market_ma']
            user_input['summary_estimated_price'] = form.cleaned_data['summary_estimated_price']
            user_input['summary_estimated_exclusive'] = form.cleaned_data['summary_estimated_exclusive']
            user_input['summary_estimated_contract'] = form.cleaned_data['summary_estimated_contract']
            user_input['summary_estimated_ea'] = form.cleaned_data['summary_estimated_ea']
            user_input['summary_estimated_em'] = form.cleaned_data['summary_estimated_em']

            #제시외건물,기계
            user_input['except_label'] = form.cleaned_data['except_label']
            user_input['except_class'] = form.cleaned_data['except_class']
            user_input['except_ho'] = form.cleaned_data['except_ho']
            user_input['except_name'] = form.cleaned_data['except_name']
            user_input['except_use'] = form.cleaned_data['except_use']
            user_input['except_size_m'] = form.cleaned_data['except_size_m']
            user_input['except_size_py'] = form.cleaned_data['except_size_py']
            user_input['except_auction_won'] = form.cleaned_data['except_auction_won']
            user_input['except_auction_m'] = form.cleaned_data['except_auction_m']
            user_input['except_auction_py'] = form.cleaned_data['except_auction_py']
            user_input['except_auction_percent'] = form.cleaned_data['except_auction_percent']
            user_input['except_market_won'] = form.cleaned_data['except_market_won']
            user_input['except_market_m'] = form.cleaned_data['except_market_m']
            user_input['except_market_py'] = form.cleaned_data['except_market_py']
            user_input['except_market_ma'] = form.cleaned_data['except_market_ma']
            user_input['except_est_won'] = form.cleaned_data['except_est_won']
            user_input['except_est_m'] = form.cleaned_data['except_est_m']
            user_input['except_est_py'] = form.cleaned_data['except_est_py']
            user_input['except_est_ea'] = form.cleaned_data['except_est_ea']
            user_input['except_est_em'] = form.cleaned_data['except_est_em']

            user_input['machine_label'] = form.cleaned_data['machine_label']
            user_input['machine_class'] = form.cleaned_data['machine_class']
            user_input['machine_ho'] = form.cleaned_data['machine_ho']
            user_input['machine_name'] = form.cleaned_data['machine_name']
            user_input['machine_use'] = form.cleaned_data['machine_use']
            user_input['machine_size_m'] = form.cleaned_data['machine_size_m']
            user_input['machine_size_py'] = form.cleaned_data['machine_size_py']
            user_input['machine_auction_won'] = form.cleaned_data['machine_auction_won']
            user_input['machine_auction_m'] = form.cleaned_data['machine_auction_m']
            user_input['machine_auction_py'] = form.cleaned_data['machine_auction_py']
            user_input['machine_auction_percent'] = form.cleaned_data['machine_auction_percent']
            user_input['machine_market_won'] = form.cleaned_data['machine_market_won']
            user_input['machine_market_m'] = form.cleaned_data['machine_market_m']
            user_input['machine_market_py'] = form.cleaned_data['machine_market_py']
            user_input['machine_market_ma'] = form.cleaned_data['machine_market_ma']
            user_input['machine_est_won'] = form.cleaned_data['machine_est_won']
            user_input['machine_est_m'] = form.cleaned_data['machine_est_m']
            user_input['machine_est_py'] = form.cleaned_data['machine_est_py']
            user_input['machine_est_ea'] = form.cleaned_data['machine_est_ea']
            user_input['machine_est_em'] = form.cleaned_data['machine_est_em']

            user_input['sum_auction_won'] = form.cleaned_data['sum_auction_won']
            user_input['sum_auction_m'] = form.cleaned_data['sum_auction_m']
            user_input['sum_auction_py'] = form.cleaned_data['sum_auction_py']
            user_input['sum_auction_percent'] = form.cleaned_data['sum_auction_percent']
            user_input['sum_market_won'] = form.cleaned_data['sum_market_won']
            user_input['sum_market_m'] = form.cleaned_data['sum_market_m']
            user_input['sum_market_py'] = form.cleaned_data['sum_market_py']
            user_input['sum_market_ma'] = form.cleaned_data['sum_market_ma']
            user_input['sum_est_won'] = form.cleaned_data['sum_est_won']
            user_input['sum_est_m'] = form.cleaned_data['sum_est_m']
            user_input['sum_est_py'] = form.cleaned_data['sum_est_py']
            user_input['sum_est_ea'] = form.cleaned_data['sum_est_ea']
            user_input['sum_est_em'] = form.cleaned_data['sum_est_em']

            #합계
            user_input['result_auction_won'] = form.cleaned_data['result_auction_won']
            user_input['result_auction_m'] = form.cleaned_data['result_auction_m']
            user_input['result_auction_py'] = form.cleaned_data['result_auction_py']
            user_input['result_auction_percent'] = form.cleaned_data['result_auction_percent']
            user_input['result_market_won'] = form.cleaned_data['result_market_won']
            user_input['result_market_m'] = form.cleaned_data['result_market_m']
            user_input['result_market_py'] = form.cleaned_data['result_market_py']
            user_input['result_market_ma'] = form.cleaned_data['result_market_ma']
            user_input['result_est_won'] = form.cleaned_data['result_est_won']
            user_input['result_est_m'] = form.cleaned_data['result_est_m']
            user_input['result_est_py'] = form.cleaned_data['result_est_py']
            user_input['result_est_ea'] = form.cleaned_data['result_est_ea']
            user_input['result_est_em'] = form.cleaned_data['result_est_em']

            #인근거래사례
            user_input['trade_up_loc'] = form.cleaned_data['trade_up_loc']
            user_input['trade_up_floor'] = form.cleaned_data['trade_up_floor']
            user_input['trade_up_structure'] = form.cleaned_data['trade_up_structure']
            user_input['trade_up_approval_date'] = form.cleaned_data['trade_up_approval_date']
            user_input['trade_up_fail_count'] = form.cleaned_data['trade_up_fail_count']
            user_input['trade_up_base_date'] = form.cleaned_data['trade_up_base_date']
            user_input['trade_up_exclusive_m'] = form.cleaned_data['trade_up_exclusive_m']
            user_input['trade_up_exclusive_py'] = form.cleaned_data['trade_up_exclusive_py']
            user_input['trade_up_right_m'] = form.cleaned_data['trade_up_right_m']
            user_input['trade_law_price_won'] = form.cleaned_data['trade_law_price_won']
            user_input['trade_law_price_exclusive'] = form.cleaned_data['trade_law_price_exclusive']
            user_input['trade_law_price_contract'] = form.cleaned_data['trade_law_price_contract']
            user_input['trade_bid_won'] = form.cleaned_data['trade_bid_won']
            user_input['trade_bid_exclusive'] = form.cleaned_data['trade_bid_exclusive']
            user_input['trade_bid_contract'] = form.cleaned_data['trade_bid_contract']
            user_input['trade_up_bidder'] = form.cleaned_data['trade_up_bidder']
            user_input['trade_up_bid_percent'] = form.cleaned_data['trade_up_bid_percent']

            user_input['trade_down_loc'] = form.cleaned_data['trade_down_loc']
            user_input['trade_down_floor'] = form.cleaned_data['trade_down_floor']
            user_input['trade_down_structure'] = form.cleaned_data['trade_down_structure']
            user_input['trade_down_approval_date'] = form.cleaned_data['trade_down_approval_date']
            user_input['trade_down_fail_count'] = form.cleaned_data['trade_down_fail_count']
            user_input['trade_down_base_date'] = form.cleaned_data['trade_down_base_date']
            user_input['trade_down_exclusive_m'] = form.cleaned_data['trade_down_exclusive_m']
            user_input['trade_down_exclusive_py'] = form.cleaned_data['trade_down_exclusive_py']
            user_input['trade_down_right_m'] = form.cleaned_data['trade_down_right_m']
            user_input['trade_down_bidder'] = form.cleaned_data['trade_down_bidder']
            user_input['trade_down_bid_percent'] = form.cleaned_data['trade_down_bid_percent']

            user_input['trade2_up_loc'] = form.cleaned_data['trade2_up_loc']
            user_input['trade2_up_floor'] = form.cleaned_data['trade2_up_floor']
            user_input['trade2_up_structure'] = form.cleaned_data['trade2_up_structure']
            user_input['trade2_up_approval_date'] = form.cleaned_data['trade2_up_approval_date']
            user_input['trade2_up_fail_count'] = form.cleaned_data['trade2_up_fail_count']
            user_input['trade2_up_base_date'] = form.cleaned_data['trade2_up_base_date']
            user_input['trade2_up_exclusive_m'] = form.cleaned_data['trade2_up_exclusive_m']
            user_input['trade2_up_exclusive_py'] = form.cleaned_data['trade2_up_exclusive_py']
            user_input['trade2_up_right_m'] = form.cleaned_data['trade2_up_right_m']
            user_input['trade2_law_price_won'] = form.cleaned_data['trade2_law_price_won']
            user_input['trade2_law_price_exclusive'] = form.cleaned_data['trade2_law_price_exclusive']
            user_input['trade2_law_price_contract'] = form.cleaned_data['trade2_law_price_contract']
            user_input['trade2_bid_won'] = form.cleaned_data['trade2_bid_won']
            user_input['trade2_bid_exclusive'] = form.cleaned_data['trade2_bid_exclusive']
            user_input['trade2_bid_contract'] = form.cleaned_data['trade2_bid_contract']
            user_input['trade2_up_bidder'] = form.cleaned_data['trade2_up_bidder']
            user_input['trade2_up_bid_percent'] = form.cleaned_data['trade2_up_bid_percent']

            user_input['trade2_down_loc'] = form.cleaned_data['trade2_down_loc']
            user_input['trade2_down_floor'] = form.cleaned_data['trade2_down_floor']
            user_input['trade2_down_structure'] = form.cleaned_data['trade2_down_structure']
            user_input['trade2_down_approval_date'] = form.cleaned_data['trade2_down_approval_date']
            user_input['trade2_down_fail_count'] = form.cleaned_data['trade2_down_fail_count']
            user_input['trade2_down_base_date'] = form.cleaned_data['trade2_down_base_date']
            user_input['trade2_down_exclusive_m'] = form.cleaned_data['trade2_down_exclusive_m']
            user_input['trade2_down_exclusive_py'] = form.cleaned_data['trade2_down_exclusive_py']
            user_input['trade2_down_right_m'] = form.cleaned_data['trade2_down_right_m']
            user_input['trade2_down_bidder'] = form.cleaned_data['trade2_down_bidder']
            user_input['trade2_down_bid_percent'] = form.cleaned_data['trade2_down_bid_percent']

            #인근낙찰사례
            user_input['bid_up_loc'] = form.cleaned_data['bid_up_loc']
            user_input['bid_up_floor'] = form.cleaned_data['bid_up_floor']
            user_input['bid_up_structure'] = form.cleaned_data['bid_up_structure']
            user_input['bid_up_approval_date'] = form.cleaned_data['bid_up_approval_date']
            user_input['bid_up_fail_count'] = form.cleaned_data['bid_up_fail_count']
            user_input['bid_up_base_date'] = form.cleaned_data['bid_up_base_date']
            user_input['bid_up_exclusive_m'] = form.cleaned_data['bid_up_exclusive_m']
            user_input['bid_up_exclusive_py'] = form.cleaned_data['bid_up_exclusive_py']
            user_input['bid_up_right_m'] = form.cleaned_data['bid_up_right_m']
            user_input['bid_law_price_won'] = form.cleaned_data['bid_law_price_won']
            user_input['bid_law_price_exclusive'] = form.cleaned_data['bid_law_price_exclusive']
            user_input['bid_law_price_contract'] = form.cleaned_data['bid_law_price_contract']
            user_input['bid_bid_won'] = form.cleaned_data['bid_bid_won']
            user_input['bid_bid_exclusive'] = form.cleaned_data['bid_bid_exclusive']
            user_input['bid_bid_contract'] = form.cleaned_data['bid_bid_contract']
            user_input['bid_up_bidder'] = form.cleaned_data['bid_up_bidder']
            user_input['bid_up_bid_percent'] = form.cleaned_data['bid_up_bid_percent']

            user_input['bid_down_loc'] = form.cleaned_data['bid_down_loc']
            user_input['bid_down_floor'] = form.cleaned_data['bid_down_floor']
            user_input['bid_down_structure'] = form.cleaned_data['bid_down_structure']
            user_input['bid_down_approval_date'] = form.cleaned_data['bid_down_approval_date']
            user_input['bid_down_fail_count'] = form.cleaned_data['bid_down_fail_count']
            user_input['bid_down_base_date'] = form.cleaned_data['bid_down_base_date']
            user_input['bid_down_exclusive_m'] = form.cleaned_data['bid_down_exclusive_m']
            user_input['bid_down_exclusive_py'] = form.cleaned_data['bid_down_exclusive_py']
            user_input['bid_down_right_m'] = form.cleaned_data['bid_down_right_m']
            user_input['bid_down_bidder'] = form.cleaned_data['bid_down_bidder']
            user_input['bid_down_bid_percent'] = form.cleaned_data['bid_down_bid_percent']

            user_input['bid2_up_loc'] = form.cleaned_data['bid2_up_loc']
            user_input['bid2_up_floor'] = form.cleaned_data['bid2_up_floor']
            user_input['bid2_up_structure'] = form.cleaned_data['bid2_up_structure']
            user_input['bid2_up_approval_date'] = form.cleaned_data['bid2_up_approval_date']
            user_input['bid2_up_fail_count'] = form.cleaned_data['bid2_up_fail_count']
            user_input['bid2_up_base_date'] = form.cleaned_data['bid2_up_base_date']
            user_input['bid2_up_exclusive_m'] = form.cleaned_data['bid2_up_exclusive_m']
            user_input['bid2_up_exclusive_py'] = form.cleaned_data['bid2_up_exclusive_py']
            user_input['bid2_up_right_m'] = form.cleaned_data['bid2_up_right_m']
            user_input['bid2_law_price_won'] = form.cleaned_data['bid2_law_price_won']
            user_input['bid2_law_price_exclusive'] = form.cleaned_data['bid2_law_price_exclusive']
            user_input['bid2_law_price_contract'] = form.cleaned_data['bid2_law_price_contract']
            user_input['bid2_bid_won'] = form.cleaned_data['bid2_bid_won']
            user_input['bid2_bid_exclusive'] = form.cleaned_data['bid2_bid_exclusive']
            user_input['bid2_bid_contract'] = form.cleaned_data['bid2_bid_contract']
            user_input['bid2_up_bidder'] = form.cleaned_data['bid2_up_bidder']
            user_input['bid2_up_bid_percent'] = form.cleaned_data['bid2_up_bid_percent']

            user_input['bid2_down_loc'] = form.cleaned_data['bid2_down_loc']
            user_input['bid2_down_floor'] = form.cleaned_data['bid2_down_floor']
            user_input['bid2_down_structure'] = form.cleaned_data['bid2_down_structure']
            user_input['bid2_down_approval_date'] = form.cleaned_data['bid2_down_approval_date']
            user_input['bid2_down_fail_count'] = form.cleaned_data['bid2_down_fail_count']
            user_input['bid2_down_base_date'] = form.cleaned_data['bid2_down_base_date']
            user_input['bid2_down_exclusive_m'] = form.cleaned_data['bid2_down_exclusive_m']
            user_input['bid2_down_exclusive_py'] = form.cleaned_data['bid2_down_exclusive_py']
            user_input['bid2_down_right_m'] = form.cleaned_data['bid2_down_right_m']
            user_input['bid2_down_bidder'] = form.cleaned_data['bid2_down_bidder']
            user_input['bid2_down_bid_percent'] = form.cleaned_data['bid2_down_bid_percent']


            #본건 거래사례, 낙찰사례
            user_input['example_name'] = form.cleaned_data['example_name']
            user_input['example_date'] = form.cleaned_data['example_date']
            user_input['example_seller'] = form.cleaned_data['example_seller']
            user_input['example_buyer'] = form.cleaned_data['example_buyer']
            user_input['example_price'] = form.cleaned_data['example_price']
            user_input['example_contract'] = form.cleaned_data['example_contract']
            user_input['example_case'] = form.cleaned_data['example_case']

            user_input['example2_name'] = form.cleaned_data['example2_name']
            user_input['example2_date'] = form.cleaned_data['example2_date']
            user_input['example2_seller'] = form.cleaned_data['example2_seller']
            user_input['example2_buyer'] = form.cleaned_data['example2_buyer']
            user_input['example2_price'] = form.cleaned_data['example2_price']
            user_input['example2_contract'] = form.cleaned_data['example2_contract']
            user_input['example2_case'] = form.cleaned_data['example2_case']

            user_input['example3_name'] = form.cleaned_data['example3_name']
            user_input['example3_date'] = form.cleaned_data['example3_date']
            user_input['example3_seller'] = form.cleaned_data['example3_seller']
            user_input['example3_buyer'] = form.cleaned_data['example3_buyer']
            user_input['example3_price'] = form.cleaned_data['example3_price']
            user_input['example3_contract'] = form.cleaned_data['example3_contract']
            user_input['example3_case'] = form.cleaned_data['example3_case']

            #종합의견
            user_input['analysis_law_price'] = form.cleaned_data['analysis_law_price']
            user_input['analysis_market_concern']=form.cleaned_data['analysis_market_concern']
            user_input['analysis_price_level'] = form.cleaned_data['analysis_price_level']
            user_input['analysis_rent_level'] = form.cleaned_data['analysis_rent_level']
            user_input['analysis_price_decision'] = form.cleaned_data['analysis_price_decision']

            #낙찰가율 통계
            user_input['statics_region'] = form.cleaned_data['statics_region']
            user_input['statics_percent_year'] = form.cleaned_data['statics_percent_year']
            user_input['statics_count_year'] = form.cleaned_data['statics_count_year']
            user_input['statics_percent_half'] = form.cleaned_data['statics_percent_half']
            user_input['statics_count_half'] = form.cleaned_data['statics_count_half']
            user_input['statics_percent_quarter'] = form.cleaned_data['statics_percent_quarter']
            user_input['statics_count_quarter'] = form.cleaned_data['statics_count_quarter']

            user_input['statics2_region'] = form.cleaned_data['statics2_region']
            user_input['statics2_percent_year'] = form.cleaned_data['statics2_percent_year']
            user_input['statics2_count_year'] = form.cleaned_data['statics2_count_year']
            user_input['statics2_percent_half'] = form.cleaned_data['statics2_percent_half']
            user_input['statics2_count_half'] = form.cleaned_data['statics2_count_half']
            user_input['statics2_percent_quarter'] = form.cleaned_data['statics2_percent_quarter']
            user_input['statics2_count_quarter'] = form.cleaned_data['statics2_count_quarter']

            user_input['statics3_region'] = form.cleaned_data['statics3_region']
            user_input['statics3_percent_year'] = form.cleaned_data['statics3_percent_year']
            user_input['statics3_count_year'] = form.cleaned_data['statics3_count_year']
            user_input['statics3_percent_half'] = form.cleaned_data['statics3_percent_half']
            user_input['statics3_count_half'] = form.cleaned_data['statics3_count_half']
            user_input['statics3_percent_quarter'] = form.cleaned_data['statics3_percent_quarter']
            user_input['statics3_count_quarter'] = form.cleaned_data['statics3_count_quarter']


            excel_write().save_file(user_input)

            BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            #MEDIA_ROOT = os.path.join(BASE_DIR, 'media')
            file_path=os.path.join(BASE_DIR, 'ibk_output.xls')
            fh=open(file_path, 'rb')
            response=HttpResponse(fh.read(), content_type='application/vnd.ms-excel')
            # 파일이름은 한글로 되어있으면 다운로드를 제공할때 올바르게 제공되지 않는다 - 이유확인해보기
            file_name=u'IBK-B-R239-01-경기도 파주시 조리읍 등원리-다세대.xls'
            final_header=file_name.encode('utf-8')

            response['Content-Disposition'] = "attachment; filename*=UTF-8\'\'%s" % django.utils.http.urlquote(file_name.encode('utf-8'))
            return response
