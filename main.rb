# -*- encoding: utf-8 -*-
require "kconv";
require './controller/GoogleApiController.rb'
require './model/SerchCondtionModel.rb'

puts "Start"
  #検索条件の設定　Start
  # パスワード
  pass = "xxxxx"
  # カレンダー情報取得アカウント
  masterAcount = "xxxxxx@xxx.co.jp"
  # アカウントリスト
  acountListMap = {"Aさん"=>"xxxxxx@xxx.co.jp", "Bさん"=>"xxxxxx@xxx.co.jp", "Cさん"=>"xxxxxx@xxx.co.jp"}
  #最大取得件数
  maxResults = 99999
  #開始日
  startMin = "2011-08-01"
  #終了日
  startMax = "2011-12-01"
  #ソート条件
  orderBy = "starttime"
  #ソート順
  sortOrder = "ascending"
  #ファイル名
  fileName = "GoogleCal出力"
  
  serchCondtionModel = SerchCondtionModel.new(pass,masterAcount,acountListMap,maxResults,startMin,startMax,orderBy,sortOrder);
  #検索条件の設定　End

  #アクションクラスオブジェクト生成
  #googleApiAction = GoogleApiAction.new("search");
  #googleApiAction = GoogleApiAction.new("excelOut");
  googleApiController = GoogleApiController.new("search");
  
  #アクションの実行
  googleApiController.doExe(serchCondtionModel,fileName);

puts "End"

