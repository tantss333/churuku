from win32com.client import Dispatch
import xlwings as xw
import pythoncom


class Function:

    def __init__(self, window):

        self.window = window



    def open_excel(self, file, orderDate):
        pythoncom.CoInitialize()

        self.file = file
        xl = xw._xlwindows.COMRetryObjectWrapper(Dispatch("Ket.Application"))
        impl = xw._xlwindows.App(visible=False, add_book=False, xl=xl)

        self.app = xw.App(visible=False, add_book=False, impl=impl)
        self.app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        self.app.screen_updating = False

        self.wb = self.app.books.open(file)
        self.sht = self.wb.sheets[-1]

        lastOrderId = str(self.sht.name)
        try:
            orderId = str(int(lastOrderId) + 1)
        except:
            orderId = ''

        try:
            orderId_original_r, orderName_base_r = self.sht.range('C4').value[::-1].split('-',1)
            orderId_original = orderId_original_r[::-1]
            orderIdName_base = orderName_base_r[::-1]

            dateTime, pcbName = orderIdName_base.split('-',1)
            orderDate = orderDate.replace('-','')

            orderName = orderDate + '-' + pcbName +'-'+ orderId
            applicant = self.sht.range('C26').value
            auditor = self.sht.range('E26').value
            approval = self.sht.range('G26').value
            steelNum = self.sht.range('C9').value

            try:
                lastOrderNum_str_r, _ = self.sht.range('E8').value[::-1].split(' ', 1)
                lastOrderNum = lastOrderNum_str_r[::-1]
            except:
                lastOrderNum = ''

            self.window.write_event_value('return-new-data',[orderName, orderId, applicant, auditor, approval, steelNum, lastOrderNum, lastOrderId])
            self.window.write_event_value('return-status','读取数据完毕')

        except Exception as e:
            txt = '解析历史数据失败，原因：%s'%e
            self.window.write_event_value('error',txt)
            self.window.write_event_value('return-status', '读取数据失败')

        finally:
            self.wb.close()
            self.app.quit()

        pythoncom.CoUninitialize()


    def close_wb(self):
        try:
            self.wb.close()
            self.app.quit()
        except:
            pass


    def add_new_sheet(self, file, orderDate, orderName, materialId, orderNum,
                                 orderId, method, pcb_mode, startDate, endDate,
                                 startNum_str, endNum_str, warrant, steelNum,
                                 bom_mode, NK_mode, NK_date, applicant, auditor, approval, SMT, DIP, whole):
        pythoncom.CoInitialize()
        try:
            xl = xw._xlwindows.COMRetryObjectWrapper(Dispatch("Ket.Application"))
            impl = xw._xlwindows.App(visible=False, add_book=False, xl=xl)
            self.app = xw.App(visible=False, add_book=False,impl=impl)
            self.app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
            self.app.screen_updating = False


            wb = self.app.books.open('./src/模板.xlsx')

        except Exception as e:
            self.window.write_event_value('error','请稍等：%s'%e)
            pythoncom.CoUninitialize()
            return


        old_sheet = wb.sheets[-1]
        lastOrderId = old_sheet.name
        old_sheet.api.Copy(Before=old_sheet.api)

        new_sheet = old_sheet
        new_sheet.name = orderId

        copied_sheet = wb.sheets[-2]  # The copied sheet is now the second last sheet
        copied_sheet.name = lastOrderId


        copied_sheet.delete()

        try:

            # Reference the original sheet and the new sheet

            new_sheet.range('B6:G6').value = [orderDate, materialId, orderNum, orderId, method, pcb_mode]
            new_sheet.range('B8:F8').value = [startDate, endDate, startNum_str, endNum_str, warrant]
            new_sheet.range('C4').value = orderName
            new_sheet.range('C26').value = applicant
            new_sheet.range('E26').value = auditor
            new_sheet.range('G26').value = approval
            new_sheet.range('C9').value = steelNum
            new_sheet.range('C10').value = NK_mode
            new_sheet.range('F9').value = bom_mode
            new_sheet.range('F10').value = NK_date
            new_sheet.range('C23').value = SMT
            new_sheet.range('C24').value = DIP
            new_sheet.range('C25').value = whole
            wb.save(file)
            self.window.write_event_value('succ','')
            self.window.write_event_value('return-status','数据保存成功')
        except Exception as e:
            txt = '保存表格数据失败，原因：%s' % e
            self.window.write_event_value('error', txt)

        finally:
            wb.close()
            try:
                self.app.quit()
            except:
                pass
        pythoncom.CoUninitialize()

    def stop_app(self):
        try:
            self.wb.close()
        except:
            pass

        try:
            self.app.quit()
        except:
            pass

