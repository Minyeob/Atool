from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField(label='업로드하실 엑셀파일을 선택해주세요')

class ExcelForm(forms.Form):

    #결과요약

    program=forms.CharField(required=False) #program
    opb=forms.CharField(required=False) #opb
    interest=forms.CharField(required=False)    #연체이자
    property_control_no=forms.CharField(required=False) #property_no
    setup_price=forms.CharField(required=False) #설정액
    user=forms.CharField(required=False)    #담당자
    user_phone=forms.CharField(required=False)  #담당자 연락처
    credit_amount=forms.CharField(required=False)   #총 채권액
    borrow_name=forms.CharField(required=False) #borrow name
    law_price=forms.CharField(required=False)   #법사가
    market_predict=forms.CharField(required=False)  #시장전망
    market_price=forms.CharField(required=False)    #시장가
    law_price_comp1=forms.CharField(required=False) #법사가 대비 1
    market_price_comp1=forms.CharField(required=False)  #시장가 대비 1
    opb_comp1=forms.CharField(required=False)   #opb 대비 1
    court=forms.CharField(required=False)   #관할법원
    bid=forms.CharField(required=False) #낙찰가
    law_price_comp2 = forms.CharField(required=False)   #법사가 대비 2
    market_price_comp2 = forms.CharField(required=False)    #시장가 대비 2
    opb_comp2 = forms.CharField(required=False) #opb 대비 2
    case=forms.CharField(required=False)    #사건번호
    avg_bid=forms.CharField(required=False) #평균낙찰가
    law_price_comp3 = forms.CharField(required=False)  # 법사가 대비 3
    market_price_comp3 = forms.CharField(required=False)  # 시장가 대비 3
    opb_comp3 = forms.CharField(required=False)  # opb 대비 3
    submission_date = forms.CharField(required=False)   #법원제출일
    next_date = forms.CharField(required=False) #차기기일
    fail_count = forms.CharField(required=False) #유찰회수

    #본건현황

    address=forms.CharField(required=False) #Address
    property_category=forms.CharField(required=False)   #Property category
    usage=forms.CharField(required=False)   #용도지역
    urlfor_lnd_info=forms.CharField(required=False)   #지목
    state=forms.CharField(required=False)   #이용상황
    land_price_m=forms.CharField(required=False)    #개별공시지가 m
    land_price_py=forms.CharField(required=False)   #개별공시지가 py
    urlfor_lnd_area=forms.CharField(required=False) #전체토지면적 m
    land_size_py=forms.CharField(required=False)   #전체토지면적 py
    security_size_m=forms.CharField(required=False) #담보면적 m
    security_size_py=forms.CharField(required=False)    #담보면적 py
    urlfor_lnd_crt=forms.CharField(required=False)   #건물 구조
    urlfor_lnd_date=forms.CharField(required=False) #사용승인일
    urlfor_lnd_mainPurps=forms.CharField(required=False) #층별 용도
    exclusive_rate=forms.CharField(required=False)  #전용율
    exclusive_area_m=forms.CharField(required=False)    #전유면적 m
    exclusive_area_py=forms.CharField(required=False)   #전유면적 py
    contract_area_m=forms.CharField(required=False) #계약면적 m
    contract_area_py=forms.CharField(required=False)    #계약면적 py

    #건물

    building_label=forms.CharField(required=False)
    building_ho=forms.CharField(required=False)
    building_exclusive_m=forms.CharField(required=False)
    building_exclusive_py=forms.CharField(required=False)
    bulding_contract_m=forms.CharField(required=False)
    building_contract_py=forms.CharField(required=False)
    building_right_m=forms.CharField(required=False)
    building_right_py=forms.CharField(required=False)
    building_ratio=forms.CharField(required=False)
    building_auction_price=forms.CharField(required=False)
    building_auction_exclusive=forms.CharField(required=False)
    building_auction_contract=forms.CharField(required=False)
    building_auction_ratio=forms.CharField(required=False)
    building_market_price = forms.CharField(required=False)
    building_market_exclusive = forms.CharField(required=False)
    building_market_contract = forms.CharField(required=False)
    bulding_market_ma=forms.CharField(required=False)
    building_estimated_price = forms.CharField(required=False)
    building_estimated_exclusive = forms.CharField(required=False)
    building_estimated_contract = forms.CharField(required=False)
    bulding_estimated_ea = forms.CharField(required=False)
    bulding_estimated_em = forms.CharField(required=False)

    summary_ho = forms.CharField(required=False)
    summary_exclusive_m = forms.CharField(required=False)
    summary_exclusive_py = forms.CharField(required=False)
    summary_contract_m = forms.CharField(required=False)
    summary_contract_py = forms.CharField(required=False)
    summary_right_m = forms.CharField(required=False)
    summary_right_py = forms.CharField(required=False)
    summary_ratio = forms.CharField(required=False)
    summary_auction_price = forms.CharField(required=False)
    summary_auction_exclusive = forms.CharField(required=False)
    summary_auction_contract = forms.CharField(required=False)
    summary_auction_ratio = forms.CharField(required=False)
    summary_market_price = forms.CharField(required=False)
    summary_market_exclusive = forms.CharField(required=False)
    summary_market_contract = forms.CharField(required=False)
    summary_market_ma = forms.CharField(required=False)
    summary_estimated_price = forms.CharField(required=False)
    summary_estimated_exclusive = forms.CharField(required=False)
    summary_estimated_contract = forms.CharField(required=False)
    summary_estimated_ea = forms.CharField(required=False)
    summary_estimated_em = forms.CharField(required=False)


    #제시외 건물, 기계기구
    except_label= forms.CharField(required=False)
    except_class= forms.CharField(required=False)
    except_ho= forms.CharField(required=False)
    except_name= forms.CharField(required=False)
    except_use= forms.CharField(required=False)
    except_size_m= forms.CharField(required=False)
    except_size_py= forms.CharField(required=False)
    except_auction_won= forms.CharField(required=False)
    except_auction_m= forms.CharField(required=False)
    except_auction_py= forms.CharField(required=False)
    except_auction_percent= forms.CharField(required=False)
    except_market_won= forms.CharField(required=False)
    except_market_m= forms.CharField(required=False)
    except_market_py= forms.CharField(required=False)
    except_market_ma= forms.CharField(required=False)
    except_est_won= forms.CharField(required=False)
    except_est_m= forms.CharField(required=False)
    except_est_py= forms.CharField(required=False)
    except_est_ea= forms.CharField(required=False)
    except_est_em= forms.CharField(required=False)

    machine_label= forms.CharField(required=False)
    machine_class= forms.CharField(required=False)
    machine_ho= forms.CharField(required=False)
    machine_name= forms.CharField(required=False)
    machine_use= forms.CharField(required=False)
    machine_size_m= forms.CharField(required=False)
    machine_size_py= forms.CharField(required=False)
    machine_auction_won= forms.CharField(required=False)
    machine_auction_m= forms.CharField(required=False)
    machine_auction_py= forms.CharField(required=False)
    machine_auction_percent= forms.CharField(required=False)
    machine_market_won= forms.CharField(required=False)
    machine_market_m= forms.CharField(required=False)
    machine_market_py= forms.CharField(required=False)
    machine_market_ma= forms.CharField(required=False)
    machine_est_won= forms.CharField(required=False)
    machine_est_m= forms.CharField(required=False)
    machine_est_py= forms.CharField(required=False)
    machine_est_ea= forms.CharField(required=False)
    machine_est_em= forms.CharField(required=False)

    sum_auction_won= forms.CharField(required=False)
    sum_auction_m= forms.CharField(required=False)
    sum_auction_py= forms.CharField(required=False)
    sum_auction_percent= forms.CharField(required=False)
    sum_market_won= forms.CharField(required=False)
    sum_market_m= forms.CharField(required=False)
    sum_market_py= forms.CharField(required=False)
    sum_market_ma= forms.CharField(required=False)
    sum_est_won= forms.CharField(required=False)
    sum_est_m= forms.CharField(required=False)
    sum_est_py= forms.CharField(required=False)
    sum_est_ea= forms.CharField(required=False)
    sum_est_em= forms.CharField(required=False)


    #합계
    result_auction_won= forms.CharField(required=False)
    result_auction_m= forms.CharField(required=False)
    result_auction_py= forms.CharField(required=False)
    result_auction_percent= forms.CharField(required=False)
    result_market_won= forms.CharField(required=False)
    result_market_m= forms.CharField(required=False)
    result_market_py= forms.CharField(required=False)
    result_market_ma= forms.CharField(required=False)
    result_est_won= forms.CharField(required=False)
    result_est_m= forms.CharField(required=False)
    result_est_py= forms.CharField(required=False)
    result_est_ea= forms.CharField(required=False)
    result_est_em= forms.CharField(required=False)

    #인근거래사례
    trade_up_loc= forms.CharField(required=False)
    trade_up_floor= forms.CharField(required=False)
    trade_up_structure= forms.CharField(required=False)
    trade_up_approval_date= forms.CharField(required=False)
    trade_up_fail_count= forms.CharField(required=False)
    trade_up_base_date= forms.CharField(required=False)
    trade_up_exclusive_m= forms.CharField(required=False)
    trade_up_exclusive_py= forms.CharField(required=False)
    trade_up_right_m= forms.CharField(required=False)
    trade_law_price_won= forms.CharField(required=False)
    trade_law_price_exclusive= forms.CharField(required=False)
    trade_law_price_contract= forms.CharField(required=False)
    trade_bid_won= forms.CharField(required=False)
    trade_bid_exclusive= forms.CharField(required=False)
    trade_bid_contract= forms.CharField(required=False)
    trade_up_bidder= forms.CharField(required=False)
    trade_up_bid_percent= forms.CharField(required=False)

    trade_down_loc= forms.CharField(required=False)
    trade_down_floor= forms.CharField(required=False)
    trade_down_structure= forms.CharField(required=False)
    trade_down_approval_date= forms.CharField(required=False)
    trade_down_fail_count= forms.CharField(required=False)
    trade_down_base_date= forms.CharField(required=False)
    trade_down_exclusive_m= forms.CharField(required=False)
    trade_down_exclusive_py= forms.CharField(required=False)
    trade_down_right_m= forms.CharField(required=False)
    trade_down_bidder= forms.CharField(required=False)
    trade_down_bid_percent= forms.CharField(required=False)

    trade2_up_loc= forms.CharField(required=False)
    trade2_up_floor= forms.CharField(required=False)
    trade2_up_structure= forms.CharField(required=False)
    trade2_up_approval_date= forms.CharField(required=False)
    trade2_up_fail_count= forms.CharField(required=False)
    trade2_up_base_date= forms.CharField(required=False)
    trade2_up_exclusive_m= forms.CharField(required=False)
    trade2_up_exclusive_py= forms.CharField(required=False)
    trade2_up_right_m= forms.CharField(required=False)
    trade2_law_price_won= forms.CharField(required=False)
    trade2_law_price_exclusive= forms.CharField(required=False)
    trade2_law_price_contract= forms.CharField(required=False)
    trade2_bid_won= forms.CharField(required=False)
    trade2_bid_exclusive= forms.CharField(required=False)
    trade2_bid_contract= forms.CharField(required=False)
    trade2_up_bidder= forms.CharField(required=False)
    trade2_up_bid_percent= forms.CharField(required=False)

    trade2_down_loc= forms.CharField(required=False)
    trade2_down_floor= forms.CharField(required=False)
    trade2_down_structure= forms.CharField(required=False)
    trade2_down_approval_date= forms.CharField(required=False)
    trade2_down_fail_count= forms.CharField(required=False)
    trade2_down_base_date= forms.CharField(required=False)
    trade2_down_exclusive_m= forms.CharField(required=False)
    trade2_down_exclusive_py= forms.CharField(required=False)
    trade2_down_right_m= forms.CharField(required=False)
    trade2_down_bidder= forms.CharField(required=False)
    trade2_down_bid_percent= forms.CharField(required=False)

    #낙찰사례
    bid_up_loc= forms.CharField(required=False)
    bid_up_floor= forms.CharField(required=False)
    bid_up_structure= forms.CharField(required=False)
    bid_up_approval_date= forms.CharField(required=False)
    bid_up_fail_count= forms.CharField(required=False)
    bid_up_base_date= forms.CharField(required=False)
    bid_up_exclusive_m= forms.CharField(required=False)
    bid_up_exclusive_py= forms.CharField(required=False)
    bid_up_right_m= forms.CharField(required=False)
    bid_law_price_won= forms.CharField(required=False)
    bid_law_price_exclusive= forms.CharField(required=False)
    bid_law_price_contract= forms.CharField(required=False)
    bid_bid_won= forms.CharField(required=False)
    bid_bid_exclusive= forms.CharField(required=False)
    bid_bid_contract= forms.CharField(required=False)
    bid_up_bidder= forms.CharField(required=False)
    bid_up_bid_percent= forms.CharField(required=False)

    bid_down_loc= forms.CharField(required=False)
    bid_down_floor= forms.CharField(required=False)
    bid_down_structure= forms.CharField(required=False)
    bid_down_approval_date= forms.CharField(required=False)
    bid_down_fail_count= forms.CharField(required=False)
    bid_down_base_date= forms.CharField(required=False)
    bid_down_exclusive_m= forms.CharField(required=False)
    bid_down_exclusive_py= forms.CharField(required=False)
    bid_down_right_m= forms.CharField(required=False)
    bid_down_bidder= forms.CharField(required=False)
    bid_down_bid_percent= forms.CharField(required=False)

    bid2_up_loc= forms.CharField(required=False)
    bid2_up_floor= forms.CharField(required=False)
    bid2_up_structure= forms.CharField(required=False)
    bid2_up_approval_date= forms.CharField(required=False)
    bid2_up_fail_count= forms.CharField(required=False)
    bid2_up_base_date= forms.CharField(required=False)
    bid2_up_exclusive_m= forms.CharField(required=False)
    bid2_up_exclusive_py= forms.CharField(required=False)
    bid2_up_right_m= forms.CharField(required=False)
    bid2_law_price_won= forms.CharField(required=False)
    bid2_law_price_exclusive= forms.CharField(required=False)
    bid2_law_price_contract= forms.CharField(required=False)
    bid2_bid_won= forms.CharField(required=False)
    bid2_bid_exclusive= forms.CharField(required=False)
    bid2_bid_contract= forms.CharField(required=False)
    bid2_up_bidder= forms.CharField(required=False)
    bid2_up_bid_percent= forms.CharField(required=False)

    bid2_down_loc= forms.CharField(required=False)
    bid2_down_floor= forms.CharField(required=False)
    bid2_down_structure= forms.CharField(required=False)
    bid2_down_approval_date= forms.CharField(required=False)
    bid2_down_fail_count= forms.CharField(required=False)
    bid2_down_base_date= forms.CharField(required=False)
    bid2_down_exclusive_m= forms.CharField(required=False)
    bid2_down_exclusive_py= forms.CharField(required=False)
    bid2_down_right_m= forms.CharField(required=False)
    bid2_down_bidder= forms.CharField(required=False)
    bid2_down_bid_percent= forms.CharField(required=False)


    #본건 거래사례,낙찰사례,유찰사례,평가전례
    example_name= forms.CharField(required=False)
    example_date= forms.CharField(required=False)
    example_seller= forms.CharField(required=False)
    example_buyer= forms.CharField(required=False)
    example_price= forms.CharField(required=False)
    example_contract= forms.CharField(required=False)
    example_case= forms.CharField(required=False)

    example2_name= forms.CharField(required=False)
    example2_date= forms.CharField(required=False)
    example2_seller= forms.CharField(required=False)
    example2_buyer= forms.CharField(required=False)
    example2_price= forms.CharField(required=False)
    example2_contract= forms.CharField(required=False)
    example2_case= forms.CharField(required=False)

    example3_name= forms.CharField(required=False)
    example3_date= forms.CharField(required=False)
    example3_seller= forms.CharField(required=False)
    example3_buyer= forms.CharField(required=False)
    example3_price= forms.CharField(required=False)
    example3_contract= forms.CharField(required=False)
    example3_case= forms.CharField(required=False)

    #종합의견
    analysis_law_price=forms.CharField(required=False)
    analysis_market_concern=forms.CharField(required=False)
    analysis_price_level = forms.CharField(required=False)
    analysis_rent_level = forms.CharField(required=False)
    analysis_price_decision = forms.CharField(required=False)

    #낙찰가율 통계
    statics_region=forms.CharField(required=False)
    statics_percent_year=forms.CharField(required=False)
    statics_count_year=forms.CharField(required=False)
    statics_percent_half=forms.CharField(required=False)
    statics_count_half=forms.CharField(required=False)
    statics_percent_quarter=forms.CharField(required=False)
    statics_count_quarter=forms.CharField(required=False)

    statics2_region=forms.CharField(required=False)
    statics2_percent_year=forms.CharField(required=False)
    statics2_count_year=forms.CharField(required=False)
    statics2_percent_half=forms.CharField(required=False)
    statics2_count_half=forms.CharField(required=False)
    statics2_percent_quarter=forms.CharField(required=False)
    statics2_count_quarter=forms.CharField(required=False)

    statics3_region=forms.CharField(required=False)
    statics3_percent_year=forms.CharField(required=False)
    statics3_count_year=forms.CharField(required=False)
    statics3_percent_half=forms.CharField(required=False)
    statics3_count_half=forms.CharField(required=False)
    statics3_percent_quarter=forms.CharField(required=False)
    statics3_count_quarter=forms.CharField(required=False)






