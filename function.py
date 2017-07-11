import xlrd
from .models import Document
import os
import xlutils.copy
import re

class excel_handling:
    #엑셀파일을 업로드하면 해당 엑셀파일을 읽어 파이썬내에서 처리할 수 있는 형태로 만드는 함수
    def make_file(self, file):
        workbook=xlrd.open_workbook(file_contents=file.read())
        return workbook

    #탁감 데이터들을 모아 선택할 수 있도록 제목을 출력해주는 함수
    def get_normal(self, workbook):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        normals=[]
        program = self.get_program_title(workbook)
        temp = program.split()
        bank = temp[0]
        for row_num in range(13, num_rows):
            type=worksheet.cell_value(row_num,2)
            code=worksheet.cell_value(row_num,1)
            pool = worksheet.cell_value(row_num, 10)
            property = worksheet.cell_value(row_num, 16)
            si_address=worksheet.cell_value(row_num,17)
            gu_address=worksheet.cell_value(row_num,18)
            dong_address=worksheet.cell_value(row_num,19)
            use=worksheet.cell_value(row_num,21)

            if(type=='탁감'):
                data=type+' '+bank+'-'+pool+'-'+property+'-'+si_address+' '+gu_address+' '+dong_address+'-'+use
                normals.append(data)

        return normals

    #탁감 데이터들의 코드를 출력하는 함수
    def get_normal_code(self, workbook):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        codes=[]
        for row_num in range(13, num_rows):
            code=worksheet.cell_value(row_num,1)
            type = worksheet.cell_value(row_num, 2)
            if (type == '탁감'):
                codes.append(code)

        return codes

    #차주명을 구해 출력해주는 함수
    def get_render_name(self, workbook, loc):
        worksheet=workbook.sheet_by_index(2)
        num_rows=worksheet.nrows
        renders=[]
        for row_num in range(13,num_rows):
            creditor=worksheet.cell_value(row_num,13)
            renders.append(creditor)
        render=renders[loc]
        return render

    #모든 형태(탁감,정밀,아파트 등)의 데이터의 코드들을 출력해주는 함수
    def get_all_code(self, workbook):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        codes = []
        for row_num in range(13, num_rows):
            code = worksheet.cell_value(row_num, 1)
            codes.append(code)

        return codes

    #해당 엑셀파일의 Program 이름을 출력해주는 함수
    def get_program_title(self, workbook):
        worksheet=workbook.sheet_by_index(0)
        program_title=worksheet.cell_value(1,0)

        return program_title

    #선택된 데이터의 property control no 를 출력해주는 함수
    def get_property_control_no(self, workbook, loc):
        worksheet=workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        program=self.get_program_title(workbook)
        temp=program.split()
        bank=temp[0]
        pools=[]
        for row_num in range(13, num_rows):
            pool = worksheet.cell_value(row_num, 10)
            pools.append(pool)
        pool=pools[loc]
        properties=[]
        for row_num in range(13, num_rows):
            property = worksheet.cell_value(row_num, 16)
            properties.append(property)
        property_code=properties[loc]
        control_no=bank+'-'+pool+'-'+property_code

        return control_no

    #각 데이터의 분류들을 모아 출력해주는 함수
    def get_type(self, workbook):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        types = []
        for row_num in range(13, num_rows):
            type = worksheet.cell_value(row_num, 2)
            types.append(type)

        return types

    #선택된 데이터의 관할법원을 출력해주는 함수
    def get_court(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        courts = []
        for row_num in range(13, num_rows):
            court = worksheet.cell_value(row_num, 72)
            courts.append(court)

        court=courts[loc]
        return court

    #선택된 데이터의 사건번호를 출력해주는 함수
    def get_case_number(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        case_numbers = []
        for row_num in range(13, num_rows):
            case = worksheet.cell_value(row_num, 73)
            case_numbers.append(case)

        case_number = case_numbers[loc]
        return case_number

    #선택된 데이터의 차주일련번호를 출력해주는 함수
    def get_render_index(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        borrower_nums = []
        for row_num in range(13, num_rows):
            borrower = worksheet.cell_value(row_num, 12)
            borrower_nums.append(borrower)

        render_index=borrower_nums[loc]
        return render_index

    #선택된 데이터의 OPB를 출력해주는 함수
    def get_opb(self, workbook, bnum):
        worksheet = workbook.sheet_by_index(0)
        num_rows = worksheet.nrows
        for row_num in range(8, num_rows):
            borrower = worksheet.cell_value(row_num, 4)
            opb=worksheet.cell_value(row_num, 13)
            if(borrower==bnum):
                result=opb

        return result

    #선택된 데이터의 연체이자를 출력해주는 함수
    def get_accured_interest(self, workbook, bnum):
        worksheet = workbook.sheet_by_index(0)
        num_rows = worksheet.nrows
        for row_num in range(8, num_rows):
            borrower = worksheet.cell_value(row_num, 4)
            interest = worksheet.cell_value(row_num, 14)
            if (borrower == bnum):
                result = interest

        return result

    #선택된 데이터의 설정액을 출력해주는 함수
    def get_cpma(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        prices = []
        for row_num in range(13, num_rows):
            setup_price = worksheet.cell_value(row_num, 30)
            prices.append(setup_price)

        cpma = prices[loc]
        return cpma

    #선택된 데이터의 총 주소를 출력해주는 함수
    def get_address(self, workbook, loc, code):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        address=[]
        ho=self.get_ho(workbook, code)
        for row_num in range(13, num_rows):
            province = worksheet.cell_value(row_num, 17)
            city = worksheet.cell_value(row_num, 18)
            district = worksheet.cell_value(row_num, 19)
            addtdistrict = worksheet.cell_value(row_num, 20)
            if(len(ho)>1):
                addtdistrict=addtdistrict+'외'

            full_address=province+' '+city+' '+district+' '+addtdistrict
            address.append(full_address)

        result=address[loc]
        return result

    #해당 건물의 용도를 구해 return 해주는 함수
    def get_property_category(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        categories = []
        for row_num in range(13, num_rows):
            use = worksheet.cell_value(row_num, 21)
            categories.append(use)

        property_category=categories[loc]
        return property_category

    #추가적인 건물의 호수들을 return 해주는 함수
    def get_ho(self, workbook, code):
        worksheet=workbook.sheet_by_index(3)
        num_rows = worksheet.nrows
        result=[]
        for row_num in range(8, num_rows):
            number=worksheet.cell_value(row_num,8)
            ho=str(worksheet.cell_value(row_num, 13))
            arr=ho.split()
            end=len(arr)

            if(number==code):
                temp = re.sub('[^0-9]','',arr[end-1])
                result.append(temp)

        return result

    #추가적인 건물의 전용적인 면적(건물면적)을 return 해주는 함수
    def get_liensize_improvement(self, workbook, code):
        worksheet = workbook.sheet_by_index(3)
        num_rows = worksheet.nrows
        result = []
        for row_num in range(8, num_rows):
            number = worksheet.cell_value(row_num, 8)
            size = worksheet.cell_value(row_num, 16)
            if (number == code):
                result.append(size)

        return result

    #추가적인 건물의 대지권의 면적을 return 해주는 함수
    def get_landsize(self, workbook, code):
        worksheet=workbook.sheet_by_index(3)
        num_rows=worksheet.nrows
        result=[]
        for row_num in range(8, num_rows):
            number=worksheet.cell_value(row_num,8)
            liensize_land=worksheet.cell_value(row_num, 15)
            land_ratio=worksheet.cell_value(row_num, 17)
            value=0
            if (number == code):
                if(land_ratio):
                    value=liensize_land*land_ratio
                result.append(value)
        return result

    #기계들의 개수가 얼마나 되는지 return 해주는 함수
    def get_utensil(self, workbook, loc):
        worksheet = workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        numbers = []
        for row_num in range(13, num_rows):
            number = worksheet.cell_value(row_num, 24)
            if(number=='상기일괄'):
                i=row_num
                while(number=='상기일괄'):
                    i=i-1
                    number=worksheet.cell_value(i,24)
                numbers.append(number)
            else:
                numbers.append(number)

        utensil = numbers[loc]
        return utensil

    #해당 사건의 주소에 대한 법정동코드를 return 해준다
    def get_address_code(self, workbook, loc):
        worksheet=workbook.sheet_by_index(2)
        num_rows = worksheet.nrows
        addresses=[]
        #엑셀파일에서 자신이 칮고자 하는 주소를 구한다
        for row_num in range(13, num_rows):
            province = worksheet.cell_value(row_num, 17)
            city = worksheet.cell_value(row_num, 18)
            district = worksheet.cell_value(row_num, 19)
            full_address = province + ' ' + city + ' ' + district
            addresses.append(full_address)

        address=addresses[loc]

        #법정동코드 엑셀파일을 연다
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        MEDIA_ROOT = os.path.join(BASE_DIR, 'media')
        file_path = os.path.join(MEDIA_ROOT, 'address_code.xlsx')
        code_workbook=xlrd.open_workbook(file_path)
        code_worksheet=code_workbook.sheet_by_index(0)
        num_rows = code_worksheet.nrows

        #내가 찾고자 하는 주소의 법정동코드를 구해 return 해준다
        for row_num in range(1,num_rows):
            code=code_worksheet.cell_value(row_num,0)
            goal=code_worksheet.cell_value(row_num,1)
            if(goal==address):
                return int(code)

    def get_addtdistrict(self, workbook, loc, code):
            worksheet = workbook.sheet_by_index(2)
            num_rows = worksheet.nrows
            address = []
            ho = self.get_ho(workbook, code)
            for row_num in range(13, num_rows):
                addtdistrict = worksheet.cell_value(row_num, 20)
                if (len(ho) > 1):
                    addtdistrict = addtdistrict + '외'
                full_address = addtdistrict
                address.append(full_address)
            result = address[loc]
            return result

    def get_adddistrict(self, workbook, loc, code):
            worksheet = workbook.sheet_by_index(2)
            num_rows = worksheet.nrows
            address = []
            ho = self.get_ho(workbook, code)
            for row_num in range(13, num_rows):
                district = worksheet.cell_value(row_num, 19)
                address.append(district)
                full_address = district
                address.append(full_address)
            result = address[loc]
            return result


class excel_write:
    def getOutCell(self, outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None

        cell = row._Row__cells.get(colIndex)
        return cell

    def setOutCell(self, outSheet, col, row, value):
        """ Change cell value without changing formatting. """
        # HACK to retain cell style.
        previousCell = self.getOutCell(outSheet, col, row)
        # END HACK, PART I

        outSheet.write(row, col, value)

        # HACK, PART II
        if previousCell:
            newCell = self.getOutCell(outSheet, col, row)
            if newCell:
                newCell.xf_idx = previousCell.xf_idx
        # END HACK

    def set_new_cell(self, outSheet, precol, prerow, col, row, value):
        previousCell = self.getOutCell(outSheet, precol, prerow)
        outSheet.write(row, col, value)

        if previousCell:
            newCell = self.getOutCell(outSheet, col, row)
            if newCell:
                newCell.xf_idx = previousCell.xf_idx


    def save_file(self, user_input):
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        MEDIA_ROOT = os.path.join(BASE_DIR, 'media')
        file_path = os.path.join(MEDIA_ROOT, 'output_sample.xls')
        inbook = xlrd.open_workbook(file_path, formatting_info=True)
        outbook = xlutils.copy.copy(inbook)
        outsheet=outbook.get_sheet(0)

        #결과요약
        self.setOutCell(outsheet, 11, 3, user_input['program'])
        self.setOutCell(outsheet, 32, 3, user_input['opb'])
        self.setOutCell(outsheet, 11, 4, user_input['user'])
        self.setOutCell(outsheet, 32, 4, user_input['interest'])
        self.setOutCell(outsheet, 11, 5, user_input['user_phone'])
        self.setOutCell(outsheet, 32, 5, user_input['credit_amount'])
        self.setOutCell(outsheet, 11, 6, user_input['property_control_no'])
        self.setOutCell(outsheet, 32, 6, user_input['setup_price'])
        self.setOutCell(outsheet, 11, 7, user_input['borrow_name'])
        self.setOutCell(outsheet, 32, 7, user_input['law_price'])
        self.setOutCell(outsheet, 11, 8, user_input['market_predict'])
        self.setOutCell(outsheet, 32, 8, user_input['market_price'])
        self.setOutCell(outsheet, 11, 9, user_input['court'])
        self.setOutCell(outsheet, 32, 9, user_input['bid'])
        self.setOutCell(outsheet, 11, 10, user_input['case'])
        self.setOutCell(outsheet, 32, 10, user_input['avg_bid'])
        self.setOutCell(outsheet, 11, 11, user_input['submission_date'])
        self.setOutCell(outsheet, 32, 11, user_input['next_date'])

        self.setOutCell(outsheet, 40, 8, user_input['law_price_comp1'])
        self.setOutCell(outsheet, 44, 8, user_input['market_price_comp1'])
        self.setOutCell(outsheet, 48, 8, user_input['opb_comp1'])
        self.setOutCell(outsheet, 40, 9, user_input['law_price_comp2'])
        self.setOutCell(outsheet, 44, 9, user_input['market_price_comp2'])
        self.setOutCell(outsheet, 48, 9, user_input['opb_comp2'])
        self.setOutCell(outsheet, 40, 10, user_input['law_price_comp3'])
        self.setOutCell(outsheet, 44, 10, user_input['market_price_comp3'])
        self.setOutCell(outsheet, 48, 10, user_input['opb_comp3'])
        self.setOutCell(outsheet, 46, 11, user_input['fail_count'])

        #본건현황
        self.setOutCell(outsheet, 55, 4, user_input['address'])
        self.setOutCell(outsheet, 91, 4, user_input['property_category'])
        self.setOutCell(outsheet, 55, 7, user_input['usage'])
        self.setOutCell(outsheet, 60, 7, user_input['urlfor_lnd_info'])
        self.setOutCell(outsheet, 64, 7, user_input['state'])
        self.setOutCell(outsheet, 69, 7, user_input['land_price_m'])
        self.setOutCell(outsheet, 75, 7, user_input['land_price_py'])
        self.setOutCell(outsheet, 82, 7, user_input['urlfor_lnd_area'])
        self.setOutCell(outsheet, 88, 7, user_input['land_size_py'])
        self.setOutCell(outsheet, 94, 7, user_input['security_size_m'])
        self.setOutCell(outsheet, 100, 7, user_input['security_size_py'])

        self.setOutCell(outsheet, 55, 10, user_input['urlfor_lnd_crt'])
        self.setOutCell(outsheet, 62, 10, user_input['urlfor_lnd_date'])
        self.setOutCell(outsheet, 69, 10, user_input['urlfor_lnd_mainPurps'])
        self.setOutCell(outsheet, 78, 10, user_input['exclusive_rate'])
        self.setOutCell(outsheet, 82, 10, user_input['exclusive_area_m'])
        self.setOutCell(outsheet, 88, 10, user_input['exclusive_area_py'])
        self.setOutCell(outsheet, 94, 10, user_input['contract_area_m'])
        self.setOutCell(outsheet, 100, 10, user_input['contract_area_py'])

        #건물
        building_count=len(user_input['building_ho'])
        height=16
        for index in range(0,building_count):
            print(user_input['building_label'][index])
            self.set_new_cell(outsheet, 2, 16, 2, height+index, user_input['building_label'][index])
            self.set_new_cell(outsheet, 5, 16, 5, height + index, user_input['building_ho'][index])
            self.set_new_cell(outsheet, 9, 16, 9, height + index, round(float(user_input['building_exclusive_m'][index]),2))
            self.set_new_cell(outsheet, 13, 16, 13, height + index, round(float(user_input['building_exclusive_py'][index]), 2))
            self.set_new_cell(outsheet, 17, 16, 17, height + index, round(float(user_input['building_contract_m'][index]), 2))
            self.set_new_cell(outsheet, 21, 16, 21, height + index, round(float(user_input['building_contract_py'][index]), 2))
            self.set_new_cell(outsheet, 25, 16, 25, height + index, round(float(user_input['building_right_m'][index]), 2))
            self.set_new_cell(outsheet, 29, 16, 29, height + index, round(float(user_input['building_right_py'][index]), 2))
            self.set_new_cell(outsheet, 33, 16, 33, height+index, user_input['building_ratio'][index])
            self.set_new_cell(outsheet, 37, 16, 37, height + index, user_input['building_auction_price'][index])
            self.set_new_cell(outsheet, 44, 16, 44, height + index, user_input['building_auction_exclusive'][index])
            self.set_new_cell(outsheet, 50, 16, 50, height + index, user_input['building_auction_contract'][index])
            self.set_new_cell(outsheet, 56, 16, 56, height + index, user_input['building_auction_ratio'][index])
            self.set_new_cell(outsheet, 59, 16, 59, height + index, user_input['building_market_price'][index])
            self.set_new_cell(outsheet, 66, 16, 66, height + index, user_input['building_market_exclusive'][index])
            self.set_new_cell(outsheet, 72, 16, 72, height + index, user_input['building_market_contract'][index])
            self.set_new_cell(outsheet, 78, 16, 78, height + index, user_input['building_market_ma'][index])
            self.set_new_cell(outsheet, 81, 16, 81, height + index, user_input['building_estimated_price'][index])
            self.set_new_cell(outsheet, 88, 16, 88, height + index, user_input['building_estimated_exclusive'][index])
            self.set_new_cell(outsheet, 94, 16, 94, height + index, user_input['building_estimated_contract'][index])
            self.set_new_cell(outsheet, 100, 16, 100, height + index, user_input['building_estimated_ea'][index])
            self.set_new_cell(outsheet, 103, 16, 103, height + index, user_input['building_estimated_em'][index])

        self.set_new_cell(outsheet, 5, 25, 5, 25, user_input['summary_ho'])
        self.set_new_cell(outsheet, 9, 25, 9, 25, user_input['summary_exclusive_m'])
        self.set_new_cell(outsheet, 13, 25, 13, 25, user_input['summary_exclusive_py'])
        self.set_new_cell(outsheet, 17, 25, 17, 25, user_input['summary_contract_m'])
        self.set_new_cell(outsheet, 21, 25, 21, 25, user_input['summary_contract_py'])
        self.set_new_cell(outsheet, 25, 25, 25, 25, user_input['summary_right_m'])
        self.set_new_cell(outsheet, 29, 25, 29, 25, user_input['summary_right_py'])
        self.set_new_cell(outsheet, 33, 25, 33, 25, user_input['summary_ratio'])
        self.set_new_cell(outsheet, 37, 25, 37, 25, user_input['summary_auction_price'])
        self.set_new_cell(outsheet, 44, 25, 44, 25, user_input['summary_auction_exclusive'])
        self.set_new_cell(outsheet, 50, 25, 50, 25, user_input['summary_auction_contract'])
        self.set_new_cell(outsheet, 56, 25, 56, 25, user_input['summary_auction_ratio'])
        self.set_new_cell(outsheet, 59, 25, 59, 25, user_input['summary_market_price'])
        self.set_new_cell(outsheet, 66, 25, 66, 25, user_input['summary_market_exclusive'])
        self.set_new_cell(outsheet, 72, 25, 72, 25, user_input['summary_market_contract'])
        self.set_new_cell(outsheet, 78, 25, 78, 25, user_input['summary_market_ma'])
        self.set_new_cell(outsheet, 81, 25, 81, 25, user_input['summary_estimated_price'])
        self.set_new_cell(outsheet, 88, 25, 88, 25, user_input['summary_estimated_exclusive'])
        self.set_new_cell(outsheet, 94, 25, 94, 25, user_input['summary_estimated_contract'])
        self.set_new_cell(outsheet, 100, 25, 100, 25, user_input['summary_estimated_ea'])
        self.set_new_cell(outsheet, 103, 25, 103, 25, user_input['summary_estimated_em'])

        #제시외건물, 기계
        self.setOutCell(outsheet, 2, 30, user_input['except_label'])
        self.setOutCell(outsheet, 5, 30, user_input['except_class'])
        self.setOutCell(outsheet, 9, 30, user_input['except_ho'])
        self.setOutCell(outsheet, 13, 30, user_input['except_name'])
        self.setOutCell(outsheet, 21, 30, user_input['except_use'])
        self.setOutCell(outsheet, 29, 30, user_input['except_size_m'])
        self.setOutCell(outsheet, 33, 30, user_input['except_size_py'])
        self.setOutCell(outsheet, 37, 30, user_input['except_auction_won'])
        self.setOutCell(outsheet, 44, 30, user_input['except_auction_m'])
        self.setOutCell(outsheet, 50, 30, user_input['except_auction_py'])
        self.setOutCell(outsheet, 56, 30, user_input['except_auction_percent'])
        self.setOutCell(outsheet, 59, 30, user_input['except_market_won'])
        self.setOutCell(outsheet, 66, 30, user_input['except_market_m'])
        self.setOutCell(outsheet, 72, 30, user_input['except_market_py'])
        self.setOutCell(outsheet, 78, 30, user_input['except_market_ma'])
        self.setOutCell(outsheet, 81, 30, user_input['except_est_won'])
        self.setOutCell(outsheet, 88, 30, user_input['except_est_m'])
        self.setOutCell(outsheet, 94, 30, user_input['except_est_py'])
        self.setOutCell(outsheet, 100, 30, user_input['except_est_ea'])
        self.setOutCell(outsheet, 103, 30, user_input['except_est_em'])

        self.setOutCell(outsheet, 2, 31, user_input['machine_label'])
        self.setOutCell(outsheet, 5, 31, user_input['machine_class'])
        self.setOutCell(outsheet, 9, 31, user_input['machine_ho'])
        self.setOutCell(outsheet, 13, 31, user_input['machine_name'])
        self.setOutCell(outsheet, 21, 31, user_input['machine_use'])
        self.setOutCell(outsheet, 29, 31, user_input['machine_size_m'])
        self.setOutCell(outsheet, 33, 31, user_input['machine_size_py'])
        self.setOutCell(outsheet, 37, 31, user_input['machine_auction_won'])
        self.setOutCell(outsheet, 44, 31, user_input['machine_auction_m'])
        self.setOutCell(outsheet, 50, 31, user_input['machine_auction_py'])
        self.setOutCell(outsheet, 56, 31, user_input['machine_auction_percent'])
        self.setOutCell(outsheet, 59, 31, user_input['machine_market_won'])
        self.setOutCell(outsheet, 66, 31, user_input['machine_market_m'])
        self.setOutCell(outsheet, 72, 31, user_input['machine_market_py'])
        self.setOutCell(outsheet, 78, 31, user_input['machine_market_ma'])
        self.setOutCell(outsheet, 81, 31, user_input['machine_est_won'])
        self.setOutCell(outsheet, 88, 31, user_input['machine_est_m'])
        self.setOutCell(outsheet, 94, 31, user_input['machine_est_py'])
        self.setOutCell(outsheet, 100, 31, user_input['machine_est_ea'])
        self.setOutCell(outsheet, 103, 31, user_input['machine_est_em'])

        self.setOutCell(outsheet, 37, 32, user_input['sum_auction_won'])
        self.setOutCell(outsheet, 44, 32, user_input['sum_auction_m'])
        self.setOutCell(outsheet, 50, 32, user_input['sum_auction_py'])
        self.setOutCell(outsheet, 56, 32, user_input['sum_auction_percent'])
        self.setOutCell(outsheet, 59, 32, user_input['sum_market_won'])
        self.setOutCell(outsheet, 66, 32, user_input['sum_market_m'])
        self.setOutCell(outsheet, 72, 32, user_input['sum_market_py'])
        self.setOutCell(outsheet, 78, 32, user_input['sum_market_ma'])
        self.setOutCell(outsheet, 81, 32, user_input['sum_est_won'])
        self.setOutCell(outsheet, 88, 32, user_input['sum_est_m'])
        self.setOutCell(outsheet, 94, 32, user_input['sum_est_py'])
        self.setOutCell(outsheet, 100, 32, user_input['sum_est_ea'])
        self.setOutCell(outsheet, 103, 32, user_input['sum_est_em'])

        #합계
        self.setOutCell(outsheet, 37, 35, user_input['result_auction_won'])
        self.setOutCell(outsheet, 44, 35, user_input['result_auction_m'])
        self.setOutCell(outsheet, 50, 35, user_input['result_auction_py'])
        self.setOutCell(outsheet, 56, 35, user_input['result_auction_percent'])
        self.setOutCell(outsheet, 59, 35, user_input['result_market_won'])
        self.setOutCell(outsheet, 66, 35, user_input['result_market_m'])
        self.setOutCell(outsheet, 72, 35, user_input['result_market_py'])
        self.setOutCell(outsheet, 78, 35, user_input['result_market_ma'])
        self.setOutCell(outsheet, 81, 35, user_input['result_est_won'])
        self.setOutCell(outsheet, 88, 35, user_input['result_est_m'])
        self.setOutCell(outsheet, 94, 35, user_input['result_est_py'])
        self.setOutCell(outsheet, 100, 35, user_input['result_est_ea'])
        self.setOutCell(outsheet, 103, 35, user_input['result_est_em'])

        #인근거래사례
        self.setOutCell(outsheet, 4, 41, user_input['trade_up_loc'])
        self.setOutCell(outsheet, 8, 41, user_input['trade_up_floor'])
        self.setOutCell(outsheet, 13, 41, user_input['trade_up_structure'])
        self.setOutCell(outsheet, 20, 41, user_input['trade_up_approval_date'])
        self.setOutCell(outsheet, 26, 41, user_input['trade_up_fail_count'])
        self.setOutCell(outsheet, 32, 41, user_input['trade_up_base_date'])
        self.setOutCell(outsheet, 38, 41, user_input['trade_up_exclusive_m'])
        self.setOutCell(outsheet, 45, 41, user_input['trade_up_exclusive_py'])
        self.setOutCell(outsheet, 51, 41, user_input['trade_up_right_m'])
        self.setOutCell(outsheet, 55, 41, user_input['trade_law_price_won'])
        self.setOutCell(outsheet, 62, 41, user_input['trade_law_price_exclusive'])
        self.setOutCell(outsheet, 68, 41, user_input['trade_law_price_contract'])
        self.setOutCell(outsheet, 74, 41, user_input['trade_bid_won'])
        self.setOutCell(outsheet, 81, 41, user_input['trade_bid_exclusive'])
        self.setOutCell(outsheet, 87, 41, user_input['trade_bid_contract'])
        self.setOutCell(outsheet, 93, 41, user_input['trade_up_bidder'])
        self.setOutCell(outsheet, 101, 41, user_input['trade_up_bid_percent'])

        self.setOutCell(outsheet, 4, 42, user_input['trade_down_loc'])
        self.setOutCell(outsheet, 8, 42, user_input['trade_down_floor'])
        self.setOutCell(outsheet, 13, 42, user_input['trade_down_structure'])
        self.setOutCell(outsheet, 20, 42, user_input['trade_down_approval_date'])
        self.setOutCell(outsheet, 26, 42, user_input['trade_down_fail_count'])
        self.setOutCell(outsheet, 32, 42, user_input['trade_down_base_date'])
        self.setOutCell(outsheet, 38, 42, user_input['trade_down_exclusive_m'])
        self.setOutCell(outsheet, 45, 42, user_input['trade_down_exclusive_py'])
        self.setOutCell(outsheet, 51, 42, user_input['trade_down_right_m'])
        self.setOutCell(outsheet, 93, 42, user_input['trade_down_bidder'])
        self.setOutCell(outsheet, 101, 42, user_input['trade_down_bid_percent'])

        self.setOutCell(outsheet, 4, 43, user_input['trade2_up_loc'])
        self.setOutCell(outsheet, 8, 43, user_input['trade2_up_floor'])
        self.setOutCell(outsheet, 13, 43, user_input['trade2_up_structure'])
        self.setOutCell(outsheet, 20, 43, user_input['trade2_up_approval_date'])
        self.setOutCell(outsheet, 26, 43, user_input['trade2_up_fail_count'])
        self.setOutCell(outsheet, 32, 43, user_input['trade2_up_base_date'])
        self.setOutCell(outsheet, 38, 43, user_input['trade2_up_exclusive_m'])
        self.setOutCell(outsheet, 45, 43, user_input['trade2_up_exclusive_py'])
        self.setOutCell(outsheet, 51, 43, user_input['trade2_up_right_m'])
        self.setOutCell(outsheet, 55, 43, user_input['trade2_law_price_won'])
        self.setOutCell(outsheet, 62, 43, user_input['trade2_law_price_exclusive'])
        self.setOutCell(outsheet, 68, 43, user_input['trade2_law_price_contract'])
        self.setOutCell(outsheet, 74, 43, user_input['trade2_bid_won'])
        self.setOutCell(outsheet, 81, 43, user_input['trade2_bid_exclusive'])
        self.setOutCell(outsheet, 87, 43, user_input['trade2_bid_contract'])
        self.setOutCell(outsheet, 93, 43, user_input['trade2_up_bidder'])
        self.setOutCell(outsheet, 101, 43, user_input['trade2_up_bid_percent'])

        self.setOutCell(outsheet, 4, 44, user_input['trade2_down_loc'])
        self.setOutCell(outsheet, 8, 44, user_input['trade2_down_floor'])
        self.setOutCell(outsheet, 13, 44, user_input['trade2_down_structure'])
        self.setOutCell(outsheet, 20, 44, user_input['trade2_down_approval_date'])
        self.setOutCell(outsheet, 26, 44, user_input['trade2_down_fail_count'])
        self.setOutCell(outsheet, 32, 44, user_input['trade2_down_base_date'])
        self.setOutCell(outsheet, 38, 44, user_input['trade2_down_exclusive_m'])
        self.setOutCell(outsheet, 45, 44, user_input['trade2_down_exclusive_py'])
        self.setOutCell(outsheet, 51, 44, user_input['trade2_down_right_m'])
        self.setOutCell(outsheet, 93, 44, user_input['trade2_down_bidder'])
        self.setOutCell(outsheet, 101, 44, user_input['trade2_down_bid_percent'])

        #인근낙찰사례
        self.setOutCell(outsheet, 4, 45, user_input['bid_up_loc'])
        self.setOutCell(outsheet, 8, 45, user_input['bid_up_floor'])
        self.setOutCell(outsheet, 13, 45, user_input['bid_up_structure'])
        self.setOutCell(outsheet, 20, 45, user_input['bid_up_approval_date'])
        self.setOutCell(outsheet, 26, 45, user_input['bid_up_fail_count'])
        self.setOutCell(outsheet, 32, 45, user_input['bid_up_base_date'])
        self.setOutCell(outsheet, 38, 45, user_input['bid_up_exclusive_m'])
        self.setOutCell(outsheet, 45, 45, user_input['bid_up_exclusive_py'])
        self.setOutCell(outsheet, 51, 45, user_input['bid_up_right_m'])
        self.setOutCell(outsheet, 55, 45, user_input['bid_law_price_won'])
        self.setOutCell(outsheet, 62, 45, user_input['bid_law_price_exclusive'])
        self.setOutCell(outsheet, 68, 45, user_input['bid_law_price_contract'])
        self.setOutCell(outsheet, 74, 45, user_input['bid_bid_won'])
        self.setOutCell(outsheet, 81, 45, user_input['bid_bid_exclusive'])
        self.setOutCell(outsheet, 87, 45, user_input['bid_bid_contract'])
        self.setOutCell(outsheet, 93, 45, user_input['bid_up_bidder'])
        self.setOutCell(outsheet, 101, 45, user_input['bid_up_bid_percent'])

        self.setOutCell(outsheet, 4, 46, user_input['bid_down_loc'])
        self.setOutCell(outsheet, 8, 46, user_input['bid_down_floor'])
        self.setOutCell(outsheet, 13, 46, user_input['bid_down_structure'])
        self.setOutCell(outsheet, 20, 46, user_input['bid_down_approval_date'])
        self.setOutCell(outsheet, 26, 46, user_input['bid_down_fail_count'])
        self.setOutCell(outsheet, 32, 46, user_input['bid_down_base_date'])
        self.setOutCell(outsheet, 38, 46, user_input['bid_down_exclusive_m'])
        self.setOutCell(outsheet, 45, 46, user_input['bid_down_exclusive_py'])
        self.setOutCell(outsheet, 51, 46, user_input['bid_down_right_m'])
        self.setOutCell(outsheet, 93, 46, user_input['bid_down_bidder'])
        self.setOutCell(outsheet, 101, 46, user_input['bid_down_bid_percent'])

        self.setOutCell(outsheet, 4, 47, user_input['bid2_up_loc'])
        self.setOutCell(outsheet, 8, 47, user_input['bid2_up_floor'])
        self.setOutCell(outsheet, 13, 47, user_input['bid2_up_structure'])
        self.setOutCell(outsheet, 20, 47, user_input['bid2_up_approval_date'])
        self.setOutCell(outsheet, 26, 47, user_input['bid2_up_fail_count'])
        self.setOutCell(outsheet, 32, 47, user_input['bid2_up_base_date'])
        self.setOutCell(outsheet, 38, 47, user_input['bid2_up_exclusive_m'])
        self.setOutCell(outsheet, 45, 47, user_input['bid2_up_exclusive_py'])
        self.setOutCell(outsheet, 51, 47, user_input['bid2_up_right_m'])
        self.setOutCell(outsheet, 55, 47, user_input['bid2_law_price_won'])
        self.setOutCell(outsheet, 62, 47, user_input['bid2_law_price_exclusive'])
        self.setOutCell(outsheet, 68, 47, user_input['bid2_law_price_contract'])
        self.setOutCell(outsheet, 74, 47, user_input['bid2_bid_won'])
        self.setOutCell(outsheet, 81, 47, user_input['bid2_bid_exclusive'])
        self.setOutCell(outsheet, 87, 47, user_input['bid2_bid_contract'])
        self.setOutCell(outsheet, 93, 47, user_input['bid2_up_bidder'])
        self.setOutCell(outsheet, 101, 47, user_input['bid2_up_bid_percent'])

        self.setOutCell(outsheet, 4, 48, user_input['bid2_down_loc'])
        self.setOutCell(outsheet, 8, 48, user_input['bid2_down_floor'])
        self.setOutCell(outsheet, 13, 48, user_input['bid2_down_structure'])
        self.setOutCell(outsheet, 20, 48, user_input['bid2_down_approval_date'])
        self.setOutCell(outsheet, 26, 48, user_input['bid2_down_fail_count'])
        self.setOutCell(outsheet, 32, 48, user_input['bid2_down_base_date'])
        self.setOutCell(outsheet, 38, 48, user_input['bid2_down_exclusive_m'])
        self.setOutCell(outsheet, 45, 48, user_input['bid2_down_exclusive_py'])
        self.setOutCell(outsheet, 51, 48, user_input['bid2_down_right_m'])
        self.setOutCell(outsheet, 93, 48, user_input['bid2_down_bidder'])
        self.setOutCell(outsheet, 101, 48, user_input['bid2_down_bid_percent'])


        #본건 거래사례
        self.setOutCell(outsheet, 1, 53, user_input['example_name'])
        self.setOutCell(outsheet, 4, 53, user_input['example_date'])
        self.setOutCell(outsheet, 10, 53, user_input['example_seller'])
        self.setOutCell(outsheet, 18, 53, user_input['example_buyer'])
        self.setOutCell(outsheet, 26, 53, user_input['example_price'])
        self.setOutCell(outsheet, 34, 53, user_input['example_contract'])
        self.setOutCell(outsheet, 41, 53, user_input['example_case'])

        self.setOutCell(outsheet, 1, 54, user_input['example2_name'])
        self.setOutCell(outsheet, 4, 54, user_input['example2_date'])
        self.setOutCell(outsheet, 10, 54, user_input['example2_seller'])
        self.setOutCell(outsheet, 18, 54, user_input['example2_buyer'])
        self.setOutCell(outsheet, 26, 54, user_input['example2_price'])
        self.setOutCell(outsheet, 34, 54, user_input['example2_contract'])
        self.setOutCell(outsheet, 41, 54, user_input['example2_case'])

        self.setOutCell(outsheet, 1, 55, user_input['example3_name'])
        self.setOutCell(outsheet, 4, 55, user_input['example3_date'])
        self.setOutCell(outsheet, 10, 55, user_input['example3_seller'])
        self.setOutCell(outsheet, 18, 55, user_input['example3_buyer'])
        self.setOutCell(outsheet, 26, 55, user_input['example3_price'])
        self.setOutCell(outsheet, 34, 55, user_input['example3_contract'])
        self.setOutCell(outsheet, 41, 55, user_input['example3_case'])

        #종합의견
        self.setOutCell(outsheet, 58, 52, user_input['analysis_law_price'])
        self.setOutCell(outsheet, 58, 54, user_input['analysis_market_concern'])
        self.setOutCell(outsheet, 58, 57, user_input['analysis_price_level'])
        self.setOutCell(outsheet, 84, 57, user_input['analysis_rent_level'])
        self.setOutCell(outsheet, 58, 60, user_input['analysis_price_decision'])

        #낙찰가율 통계
        self.setOutCell(outsheet, 1, 60, user_input['property_category'])
        self.setOutCell(outsheet, 9, 60, user_input['statics_region'])
        self.setOutCell(outsheet, 16, 60, user_input['statics_percent_year'])
        self.setOutCell(outsheet, 23, 60, user_input['statics_count_year'])
        self.setOutCell(outsheet, 28, 60, user_input['statics_percent_half'])
        self.setOutCell(outsheet, 35, 60, user_input['statics_count_half'])
        self.setOutCell(outsheet, 40, 60, user_input['statics_percent_quarter'])
        self.setOutCell(outsheet, 47, 60, user_input['statics_count_quarter'])

        self.setOutCell(outsheet, 9, 61, user_input['statics2_region'])
        self.setOutCell(outsheet, 16, 61, user_input['statics2_percent_year'])
        self.setOutCell(outsheet, 23, 61, user_input['statics2_count_year'])
        self.setOutCell(outsheet, 28, 61, user_input['statics2_percent_half'])
        self.setOutCell(outsheet, 35, 61, user_input['statics2_count_half'])
        self.setOutCell(outsheet, 40, 61, user_input['statics2_percent_quarter'])
        self.setOutCell(outsheet, 47, 61, user_input['statics2_count_quarter'])

        self.setOutCell(outsheet, 9, 62, user_input['statics3_region'])
        self.setOutCell(outsheet, 16, 62, user_input['statics3_percent_year'])
        self.setOutCell(outsheet, 23, 62, user_input['statics3_count_year'])
        self.setOutCell(outsheet, 28, 62, user_input['statics3_percent_half'])
        self.setOutCell(outsheet, 35, 62, user_input['statics3_count_half'])
        self.setOutCell(outsheet, 40, 62, user_input['statics3_percent_quarter'])
        self.setOutCell(outsheet, 47, 62, user_input['statics3_count_quarter'])





        outbook.save('ibk_output.xls')
        #new_document = Document(file=os.path.join(BASE_DIR, 'output.xls'))
        #new_document.title = 'ibk_output.xls'
        #file = Document.objects.filter(title=new_document.title)
        #if (len(file) == 0):
        #   new_document.save()
        return os.path.join(BASE_DIR, 'ibk_output.xls')