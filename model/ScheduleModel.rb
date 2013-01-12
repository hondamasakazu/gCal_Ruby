# -*- encoding: utf-8 -*-
require "gcalapi";

class ScheduleModel
  # コンストラクタ。
  def initialize(googleCalObj)
    # GoogleCalendarを受け取り、クラス変数に保持する。
    @googleCalObj = googleCalObj;
  end
  attr_accessor :googleCalObj
  
  
  # メソッド定義。
  def getTitle()
    return googleCalObj.title
  end
  def getDesc()
    return googleCalObj.desc
  end
  def getWhere()
    return googleCalObj.en
  end
  def getEndDate2()
    return googleCalObj.en
  end
  def getStartDate()
    return googleCalObj.st
  end
  def getEndDate()
    return googleCalObj.en
  end
end
