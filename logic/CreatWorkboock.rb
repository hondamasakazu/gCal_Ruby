# -*- encoding: utf-8 -*-
require './model/SerchCondtionModel.rb'
require './model/SerchOutputModel.rb'
require './logic/SerchGoogleCalender.rb'
require './common/GcalUtils.rb'
require 'spreadsheet'

=begin 
Excel生成/出力クラス
=end
class CreatWorkboock

  TITLE_INDEX_ROW  = 4  
  TITLE_DATA_COL   = 2
  TITLE_TITLE_COL  = 3
  TITLE_DETAIL_COL = 4
  TITLE_KOSU_COL   = 5

  def initialize(serchOutputModelList,finame)
    @serchOutputModelList = serchOutputModelList
    @finame = finame
  end

=begin 
Excel生成/出力を実行
=end
  def doExe() 
    Spreadsheet.client_encoding = 'UTF-8'
    workbook = Spreadsheet::Workbook.new
    
    @serchOutputModelList.each do |serchOutputModel|
        creatWorksheet(workbook, serchOutputModel)
    end
    #ファイル名を指定しファイル出力実行
    workbook.write(@finame + '.xls')
  end

=begin 
Excelシートを生成
=end
  def creatWorksheet(workbook, serchOutputModel)
    
    #シートを生成
    worksheet = workbook.create_worksheet(:name => serchOutputModel.name)
    worksheet.default_format = ExcelFormatController::NormalFormat.new
    
    #ヘッダー行を生成
    worksheet[2, 2] = "アカウント"
    worksheet.row(2).set_format(2, ExcelFormatController::TitleDefaultFormat.new)
    worksheet[2, 3] = serchOutputModel.acount
    worksheet.row(2).set_format(3, ExcelFormatController::LabelDefaultFormat.new)
    
    #タイトル行を生成
    self.creatTitleLine(worksheet);
    
    #ラベル行を生成
    indexRow = TITLE_INDEX_ROW + 1;
    totalKosu = 0;
    dateUtils = DateController::DateUtils.new
    serchOutputModel.eventList.each do |event|
      self.creatLabelLine(worksheet,indexRow, event);
      totalKosu = totalKosu + dateUtils.calcTime(event.st,event.en)
      indexRow = indexRow + 1;
    end
    
    #フッター行を生成
    worksheet[indexRow + 1, TITLE_DETAIL_COL] = "合計"
    worksheet.row(indexRow + 1).set_format(TITLE_DETAIL_COL, ExcelFormatController::TitleDefaultFormat.new)
    worksheet[indexRow + 1, TITLE_KOSU_COL] =  totalKosu
    worksheet.row(indexRow + 1).set_format(TITLE_KOSU_COL, ExcelFormatController::LabelDefaultFormat.new)
    
  end

=begin 
タイトル行を生成
=end
  def creatTitleLine(worksheet)
    titleDefaultFormat = ExcelFormatController::TitleDefaultFormat.new
    worksheet[TITLE_INDEX_ROW, TITLE_DATA_COL] = "日時"
    worksheet.row(TITLE_INDEX_ROW).set_format(TITLE_DATA_COL, titleDefaultFormat)
    worksheet[TITLE_INDEX_ROW, TITLE_TITLE_COL] = "タイトル"
    worksheet.row(TITLE_INDEX_ROW).set_format(TITLE_TITLE_COL, titleDefaultFormat)
    worksheet[TITLE_INDEX_ROW, TITLE_DETAIL_COL] = "詳細"
    worksheet.row(TITLE_INDEX_ROW).set_format(TITLE_DETAIL_COL, titleDefaultFormat)
    worksheet[TITLE_INDEX_ROW, TITLE_KOSU_COL] = "工数(H)"
    worksheet.row(TITLE_INDEX_ROW).set_format(TITLE_KOSU_COL, titleDefaultFormat)
  end

=begin 
ラベル行を生成
=end
  def creatLabelLine(worksheet,indexRow, event)
    labelDefaultFormat = ExcelFormatController::LabelDefaultFormat.new
    #日時
    dateUtils = DateController::DateUtils.new
    worksheet[indexRow, TITLE_DATA_COL] = dateUtils.dataFormat(dateUtils.conversionJST(event.st)) + 
                                          "～" + 
                                          dateUtils.dataFormat(dateUtils.conversionJST(event.en))
    worksheet.row(indexRow).set_format(TITLE_DATA_COL, labelDefaultFormat)
    #タイトル
    worksheet[indexRow, TITLE_TITLE_COL] = event.title
    worksheet.row(indexRow).set_format(TITLE_TITLE_COL, labelDefaultFormat)
    #詳細
    worksheet[indexRow, TITLE_DETAIL_COL] = event.desc
    worksheet.row(indexRow).set_format(TITLE_DETAIL_COL, labelDefaultFormat)
    #工数
    worksheet[indexRow, TITLE_KOSU_COL] = dateUtils.calcTime(event.st,event.en)
    worksheet.row(indexRow).set_format(TITLE_KOSU_COL, labelDefaultFormat)
  end

end