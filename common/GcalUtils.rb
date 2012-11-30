# -*- encoding: utf-8 -*-

=begin 
共通ロジックモジュール群を定義する
=end

=begin 
Data型オブジェクトを操作する共通モジュールを定義する
=end
module DateController
require 'date'
require 'time'

=begin 
Data型オブジェクトを操作する
=end
  class DateUtils
    def initialize()
    end

=begin 
引数で渡された日時の差分を返却します。（単位：時間）
=end
    def calcTime(st,en)
      starTime = DateTime.new(st.year, st.mon, st.day, st.hour, st.min, st.sec)
      endTime = DateTime.new(en.year, en.mon, en.day, en.hour, en.min, en.sec)
      return (endTime - starTime) * 24
    end

=begin 
引数で渡された日時に対してプラス９時間します（UTC　⇒　JSTへ変換する）
=end
    def conversionJST(dateTime)
      t = DateTime.new(dateTime.year, dateTime.mon, dateTime.day, dateTime.hour, dateTime.min)
      #9時間プラス
      t = t + Rational(9, 24)
      return t
    end
    
    def dataFormat(dateTime)
      dateTime.strftime "%Y/%m/%d %H:%M"
    end
    
  end

end

=begin 
Excelオブジェクトを操作する共通モジュールを定義する
=end
module ExcelFormatController
require 'spreadsheet'

=begin 
NormalFormat

  size:11
=end
  class NormalFormat < Spreadsheet::Format
    def initialize(opts={})
      super(opts.merge(:size => 11))
    end
  end

=begin 
TitleDefaultFormat  

  color            :black 
  weight           :bold
  horizontal_align :center
  pattern_fg_color :yellow
=end
  class TitleDefaultFormat < NormalFormat
    def initialize
      self.border = true
      super(:color            => :black, 
            :weight           => :bold,
            :horizontal_align => :center,
            :pattern => 1,
            :pattern_fg_color => :yellow)
    end
  end

=begin
LabelDefaultFormat

  border           :true
  color            :black
  horizontal_align :left
=end
  class LabelDefaultFormat < NormalFormat
    def initialize
      self.border = true
      super(:color            => :black,
            :horizontal_align => :left)
    end
  end
end