# -*- encoding: utf-8 -*-
require "kconv";
require "gcalapi";
require './model/ScheduleModel.rb'
require './model/SerchCondtionModel.rb'
require './model/SerchOutputModel.rb'
require './logic/SerchGoogleCalender.rb'
require './logic/CreatWorkboock.rb'

=begin 
Googleカレンダー情報コントロールアクションクラス  
=end
class GoogleApiController

  def initialize(cmd)
    @cmd = cmd
  end

=begin 
Googleカレンダー情報取得実行
=end
  def doExe(serchCondtionModel, fileName)
    #カレンダー情報取得クラスオブジェクト生成
    serchGoogleCalender = SerchGoogleCalender.new(serchCondtionModel);
    #カレンダー情報取得
    serchOutputModelList = serchGoogleCalender.getGoogleCalenderData();

    #アクションが検索
    if @cmd == "search" then
      #デバック用（本来は画面表示用のオブジェクトへ詰め替える予定）
      cnt = 1;
      serchOutputModelList.each do |serchOutputModel|
        serchOutputModel.eventList.each do |event|
          puts serchOutputModel.name
          puts serchOutputModel.acount
          #タイトル
          #puts event.title
          #詳細
          #puts event.desc
          #場所
          #puts event.where
          #開始日時
          #puts event.st
          #終了日時
          #puts event.en
          
          schejule = ScheduleModel.new(event);
          puts schejule.getTitle()
          puts schejule.getDesc()
          puts schejule.getWhere()
          puts schejule.getStartDate()
          puts schejule.getEndDate()
          
          print(cnt , "件目")
          cnt = cnt + 1
        end
      end
    #アクションがExcel出力
    elsif @cmd == "excelOut" then
      #Excelオブジェクトクラス生成
      creatWorkboock = CreatWorkboock.new(serchOutputModelList, fileName);
      #Excel出力実行
      creatWorkboock.doExe;
    end

  end
end
