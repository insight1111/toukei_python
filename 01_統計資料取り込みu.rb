#
#  毎月の月次経営動態資料からデータを取り込む
#  必ず月次経営動態を開いた後に実行すること
#

#エラーが発生したら、それはerr.txtに保存する
#$stderr=File.open("c:/err.txt","w")
require 'win32ole'
require 'sqlite3'
require 'kconv'
require 'yaml'
require 'pp'
# require 'vr/vruby'
#require 'ruby-debug'


#データベースに書き込むための関数（関数中のコメントは重複チェック。
#かなり遅くなるので、今はコメントアウトしている
def write_database(db,nendo,year,month,ka_id,cont,value,memo,nyugai,view)
  # if @db.execute("select * from "+db+" where nendo=? and month=? and ka_id=? and cont=? and nyugai=?",nendo,month,ka_id,cont,nyugai).to_ary.size==0
  sql="insert into "+db+" (nendo,year,month,ka_id,cont,value,memo,nyugai,view) values(?,?,?,?,?,?,?,?,?)"
  @db.execute(sql,nendo,year,month,ka_id,cont.toutf8, value,memo,nyugai,view)
  # else
  #  sql="update "+db+" set value=? where nendo=? and month=? and ka_id=? and cont=? and nyugai=?"
  #    @db.execute(sql,value,nendo,month,ka_id,cont.toutf8,nyugai)
  # end
end

#database.iniを調べて、使用データベースを決定する
begin
  open("database.ini") do |file|
    while f=file.gets
      next if f=~/^#/
        Database=f.chomp!
    end
  end
rescue
  wsh = WIN32OLE.new('WScript.Shell')
  wsh.Popup("database.iniのファイルにデータベースの場所を書き込んで下さい",0, "database.iniが読み込めません")
  open("database.ini","w") do |file|
    file.puts "#データベースの場所を書き込んで下さい。"
  end
  exit
end

#フォーマット設定ファイルの読み込み
begin
  d=YAML.load_file('toukei_format.ini')
rescue
  wsh = WIN32OLE.new('WScript.Shell')
  wsh.Popup("toukei_format.iniのファイルにデータベースの場所を書き込んで下さい",0, "toukei_format.iniが読み込めません")
end
def cp(o)
  # debugger
  o.inject({}) do |hash, (k, v)|
    case v
    when Hash
      hash[k.encode('cp932')]=cp(v)
    else
      hash[k.encode('cp932')]=v
    end
  hash
  end
end
# p d
# d=cp(d)
# p d
begin
  print "年月を入力してください（例：2008年1月→200801）"
  ym=gets.chomp
  raise if ym.size!=6
  year=ym[0..3].to_i
  month=ym[4,5].to_i
  nendo=year
  nendo-=1 if month>=1 && month <=3
  raise if month<1 || month>12
rescue
  retry
end
# nendo=2021
# year=2021
# month=5

@db=SQLite3::Database.new(Database)

# sql="select sum(value) from data where nendo=#{nendo} and month=#{month} and cont='"+"合計".toutf8+"' and nyugai=0"
# if re=@db.execute(sql)[0][0]
#   puts "既にデータは取り込まれています"
#   puts "Enterキーを押してください..."
#   gets
#   exit
# end
wb=WIN32OLE.connect("Excel.Application").activeworkbook

begin

@db.transaction do
  sh=wb.sheets("#{month}月")

  #sql="insert into data (nendo,year,month,ka_id,cont,value,memo,nyugai,view) values(?,?,?,?,?,?,?,?,?)"
  datafield={}

  #確定項目parse
  puts "項目調査中"
  td=d["確定"]
  (td["左上"]["行"]+1..td["総合計"]["行"]).each do |row|
    temp=sh.cells(row,td["左上"]["列"]).value
    #puts datafield.to_s.toutf8
    next if temp==nil or temp=="小計"
    #para row
    datafield[row-td["オフセット"]]=temp
  end
  td["その他"].each do |temp|
    datafield[temp[0]-td["オフセット"]]=temp[1]
  end


  #確定dataインポート
  #[科番号、入外（入=0,外=1)]
  puts "確定"
  ka=[
    [100, 0, td["全体入院列"]],
    [100, 1, td["全体外来列"]],
    [ 50, 1, td["呉服町列"]],
    [ 75, 1, td["泌クリニック列"]]
  ]
  ka.each do |ka|
   datafield.each do |row,cont|
  		#puts row
     temp_data=sh.cells(row+td["オフセット"],ka[2]).value.to_i
     write_database("data",nendo,year,month,ka[0],cont.toutf8, temp_data,nil,ka[1],row)
   end
  end

  #病院全体用dataインポート
  td=d["病院全体"]
  dd=td["延べ入院患者数"]
  write_database("data",nendo,year,month,100,"延べ入院患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["新入院患者数"]
  write_database("data",nendo,year,month,100,"新入院患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["退院患者数"]
  write_database("data",nendo,year,month,100,"退院患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["手術件数"]
  write_database("data",nendo,year,month,100,"手術件数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["ESWL"]
  #debugger
  write_database("data",nendo,year,month,100,"ESWL件数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["DS件数"]
  write_database("data",nendo,year,month,100,"DS件数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,0,0)
  dd=td["延べ外来患者数"]
  write_database("data",nendo,year,month,100,"延べ外来患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,91,1,0)
  dd=td["外来初診患者数"]
  dd2=d["歯科"]["外来初診患者数"]
  # write_database("data",nendo,year,month,100,"外来初診患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i+sh.cells(dd2["行"],dd2["列"]).value.to_i,92,1,0)
  write_database("data",nendo,year,month,100,"外来初診患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,92,1,0)
  td2=td["新規登録患者数"]
  dd=td2["全体"]
  dd1=td2["歯科"]
  dd2=td2["健診"]
  dd3=td2["クリニック"]
  write_database("data",nendo,year,month,100,"新規登録患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i-sh.cells(dd1["行"],dd1["列"]).value.to_i-sh.cells(dd2["行"],dd2["列"]).value.to_i-sh.cells(dd3["行"],dd3["列"]).value.to_i,93,1,0)
  dd=td["平均在院日数"]
  write_database("data",nendo,year,month,100,"平均在院日数".toutf8, sh.cells(dd["行"],dd["列"]).value,nil,0,0)
  sql="update data set value=? where nendo=? and month=? and ka_id=? and cont=? and nyugai=?"
  @db.execute(sql,sh.cells(d["確定"]["総合計"]["行"],d["確定"]["総合計"]["列"]).value.to_i,nendo,month,100,"合計".toutf8, 1)

  #呉服町(75)、泌クリ(50)
  dd=d["呉服町"]["延べ外来患者数"]
  write_database("data",nendo,year,month,50,"延べ外来患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,1,0)
  dd=d["泌クリニック"]["延べ外来患者数"]
  write_database("data",nendo,year,month,75,"延べ外来患者数".toutf8, sh.cells(dd["行"],dd["列"]).value.to_i,nil,1,0)

  ka={"内科"=>1,"外科"=>3,"乳腺内分泌外科"=>29,"整形外科"=>4,"泌尿器科"=>5,"婦人科"=>7,"放射線科"=>12,"歯科"=>10,
  "血液内科"=>16,"消化管内科"=>17,"循環器科"=>18,"脳神経外科"=>19,
  "腎臓内科"=>22,"肝胆膵内科"=>24,"総合診療科"=>41,"脳神経内科"=>42,"呼吸器科"=>51,
  "糖尿病科"=>71,"睡眠呼吸障害センター"=>81,"内科リウマチ科"=>82,
    "リウマチ科"=>82,"内科・リウマチ科"=>82,"神経内科"=>42,
    "睡眠時無呼吸センター"=>81, "SAS"=>81}
    p ka["外科"]

  #医師別インポート
  puts "医師別"
  td=d["医師別"]
  reject=%w(神経科 歯科 精神科 皮膚科 眼科 耳鼻科 不妊 産科 透析科 感染症免疫科 内科リウマチ科 健診科 麻酔科 .)
  reject << ""
  start_line=td["入院"]["左上"]["行"]
  sabun=td["外来"]["左上"]["列"]-td["入院"]["左上"]["列"]
  (td["入院"]["左上"]["列"]+1..td["入院"]["右下"]["列"]-3).each do |col|
    next if reject.index(sh.cells(start_line,col).value)
    datafield.each do |row,val|
	#puts row
      nyuin_data=sh.cells(row+start_line,col).value.to_i
      gairai_data=sh.cells(row+start_line,col+sabun).value.to_i
      #puts "cellka= #{sh.cells(start_line,col).value} ka=#{ka[sh.cells(start_line,col).value]}"
      #puts "encoding= #{sh.cells(start_line,col).value.encoding}"
      write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value.split(/[:：]/)[0]],val, nyuin_data,nil,0,row) if val != "合計"
      # write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value],val.toutf8, nyuin_data,nil,0,row) if val != "合計"
      write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value.split(/[:：]/)[0]],val, gairai_data,nil,1,row) if val != "合計"
      # write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value],val.toutf8, gairai_data,nil,1,row) if val != "合計"
      # write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value.split(/[:：]/)[0]],"延べ入院患者数".toutf8, nyuin_data,nil,0,row) if row+start_line==td["診療日数列"]
      write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value.split(/[:：]/)[0]],"延べ外来患者数".toutf8, gairai_data,nil,1,row) if row+start_line==td["診療日数列"]
    end
    puts ka[sh.cells(start_line,col).value]
    write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value],"合計".toutf8, sh.cells(td["合計列"],col).value,nil,0,0)
    write_database("data",nendo,year,month,ka[sh.cells(start_line,col).value],"合計".toutf8, sh.cells(td["合計列"],col+sabun).value,nil,1,0)
  end
  
  #入院患者数
  td=d["入院患者数"]
  puts "入院患者数"
  # puts td["開始行"]
  # puts td["終了行"]
  # puts td["列"]
  (td["開始行"]..td["終了行"]).each do |row|
    data=sh.cells(row,td["列"]).value.to_i
    write_database("data",nendo,year,month,ka[sh.cells(row,10).value.split(/[:：]/)[0]],"延べ入院患者数", data,nil,0,row)
  end
  #診療日数インポート
  #puts "診療日数"
  #write_database("data",nendo,year,month,100,"診療実日数".toutf8,sh.cells(2,12).value, nil,0,1)
  #write_database("data",nendo,year,month,100,"診療実日数".toutf8,sh.cells(3,12).value, nil,1,1)

  #平均在院日数
  #  puts "平均在院日数"
  #  r=d[""]["延べ外来患者数"]["行"]
  #  c=d["呉服町"]["延べ外来患者数"]["列"]
  #  write_database("data",nendo,year,month,nil,"平均在院日数".toutf8,sh.cells(5,11).value, nil,99,1)

  #患者数統計
  puts "患者数統計"
  td=d["患者数"]
  (td["入院"]["左上"]["行"]+1..td["入院"]["右下"]["行"]).each do |row|
    write_database("data",nendo,year,month,nil,sh.cells(row,td["入院"]["左上"]["列"]).value.toutf8,sh.cells(row,td["入院"]["左上"]["列"]+1).value, nil,0,0)
    write_database("data",nendo,year,month,nil,sh.cells(row,td["外来"]["左上"]["列"]).value.toutf8,sh.cells(row,td["外来"]["左上"]["列"]+1).value, nil,1,0)
  end

  #科別入院患者数
  puts "科別入院患者数"
  td=d["科別入院患者数"]
  (td["先頭行"]+1..td["最終行"]-1).each do |row|
    p row
    # write_database("data",nendo,year,month,ka[sh.cells(row,td["入院列"]-1).value],"新入院患者数",sh.cells(row,td["入院列"]).value, nil,0,0)
    # write_database("data",nendo,year,month,ka[sh.cells(row,td["入院列"]-1).value],"退院患者数",sh.cells(row,td["退院列"]).value, nil,0,0)
    write_database("data",nendo,year,month,ka[sh.cells(row,td["入院列"]-1).value.split(/[:：]/)[0]],"新入院患者数".toutf8,sh.cells(row,td["入院列"]).value, nil,0,0)
    write_database("data",nendo,year,month,ka[sh.cells(row,td["入院列"]-1).value.split(/[:：]/)[0]],"退院患者数".toutf8,sh.cells(row,td["退院列"]).value, nil,0,0)
    write_database("data",nendo,year,month,ka[sh.cells(row,10).value],"延べ入院患者数".toutf8,sh.cells(row,13).value, nil,0,0)
  end

  #死亡患者数
  puts "死亡患者数"
  td=d["死亡患者数"]
  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
    write_database("data",nendo,year,month,ka[sh.cells(row,td["左上"]["列"]).value],"死亡患者数".toutf8,sh.cells(row,td["左上"]["列"]+1).value, nil,0,0)
  end

  #病理解剖件数
  puts "病理解剖件数"
  td=d["病理解剖数"]
  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
    write_database("data",nendo,year,month,ka[sh.cells(row,td["左上"]["列"]).value],"病理解剖件数".toutf8,sh.cells(row,td["右下"]["列"]).value, nil,0,0)
  end

  #外来新患者数
  puts "外来新患者数"
  td=d["外来新患者数"]
  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
  next if sh.cells(row,td["左上"]["列"])=="歯科：10"
    write_database("data",nendo,year,month,sh.cells(row,td["左上"]["列"]).value.split(/[：:]/)[1].to_i,"外来初診患者数".toutf8,sh.cells(row,td["右下"]["列"]).value, nil,1,0)
    # write_database("data",nendo,year,month,sh.cells(row,td["左上"]["列"]).value.split(/[:：]/)[1].to_i,"外来初診患者数".toutf8,sh.cells(row,td["右下"]["列"]).value, nil,1,0)
  end

  #紹介・救急搬送件数
  #  puts "紹介・救急搬送件数"
  #  (54..68).each do |row|
  #    write_database("data"2,nendo,year,month,sh.cells(row,10).value.split(/[:：]/)[1].to_i,"紹介患者数".toutf8,sh.cells(row,14).value, nil,99,0)
  #    write_database("data"2,nendo,year,month,sh.cells(row,10).value.split(/[:：]/)[1].to_i,"救急車件数".toutf8,sh.cells(row,15).value, nil,99,0)
  #  end
  #    write_database("data"2,nendo,year,month,100,"紹介患者数".toutf8,sh.cells(69,14).value, nil,99,0)
  #    write_database("data"2,nendo,year,month,100,"救急車件数".toutf8,sh.cells(69,15).value, nil,99,0)

  #新規登録患者数
  puts "新規登録患者数"
  td=d["新規登録患者数"]
  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
    write_database("data",nendo,year,month,sh.cells(row,td["左上"]["列"]).value.split(/[：:]/)[1].to_i,"新規登録患者数".toutf8,sh.cells(row,td["右下"]["列"]).value, nil,1,0)
  end

  #手術件数
  puts "手術件数"
  td=d["手術件数"]
  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
    write_database("data",nendo,year,month,sh.cells(row,td["左上"]["列"]).value.split(/[：:]/)[1].to_i,"手術件数".toutf8,sh.cells(row,td["左上"]["列"]+1).value, nil,0,0)
      write_database("data",nendo,year,month,sh.cells(row,td["左上"]["列"]).value.split(/[：:]/)[1].to_i,"DS件数".toutf8,sh.cells(row,td["左上"]["列"]+2).value, nil,0,0)
  end

  #歯科
  puts "歯科"
  td=d["歯科"]
  dd=td["合計"]
  write_database("data",nendo,year,month,10,"合計".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)
  dd=td["外来初診患者数"]
  #write_database("data",nendo,year,month,10,"外来初診患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)
  dd2=td["延べ外来患者数"]
  write_database("data",nendo,year,month,10,"延べ外来患者数".toutf8,sh.cells(dd2["行"],dd2["列"]).value.to_i+sh.cells(dd["行"],dd["列"]).value.to_i, nil,1,0)
  dd=td["入院収入"]
  puts dd["行"]+dd["列"]
  write_database("data",nendo,year,month,10,"合計".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  dd=td["新入院患者数"]
  #write_database("data",nendo,year,month,10,"新入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  dd=td["延べ入院患者数"]
  #write_database("data",nendo,year,month,10,"延べ入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)

  #健診
  puts "健診"
  td=d["健診"]
  dd=td["入院収入"]
  write_database("data",nendo,year,month,90,"合計".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  dd=td["外来収入"]
  write_database("data",nendo,year,month,90,"合計".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)
  dd=td["新入院患者数"]
  write_database("data",nendo,year,month,90,"新入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  dd=td["延べ入院患者数"]
  write_database("data",nendo,year,month,90,"延べ入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  dd=td["延べ外来患者数"]
  write_database("data",nendo,year,month,90,"延べ外来患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)
  write_database("data",nendo,year,month,90,"外来初診患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)

  # puts "歯科"
  # td=d["歯科"]
  # dd=td["新入院患者数"]
  # write_database("data",nendo,year,month,10,"新入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  # # dd=td["延べ入院患者数"]
  # # write_database("data",nendo,year,month,10,"延べ入院患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  # dd=td["手術件数"]
  # write_database("data",nendo,year,month,10,"手術件数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,0,0)
  # dd=td["外来初診患者数"]
  # p dd["行"]
  # p dd["列"]
  # write_database("data",nendo,year,month,10,"外来初診患者数".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)

  #おおはま
  puts "おおはま"
  dd=d["おおはま"]["収入"]
  write_database("data",nendo,year,month,97,"合計".toutf8,sh.cells(dd["行"],dd["列"]).value, nil,1,0)

  #病棟患者数
  puts "病棟患者数"
  td=d["病棟患者数"]
  byoto={ "２－Ⅰ"=>200, "２－Ⅱ"=>201, "３－Ⅰ"=>202, "３－Ⅱ"=>203, "４－Ⅰ"=>204, "４－婦"=>205, "４－Ⅱ"=>206, "５－Ⅱ"=>207, "５－Ⅰ"=>208, "６Ｆ"=>209, "５－Ⅰ、６Ｆ"=>210 }
  temp_list=[[td["新入院列"],"新入院"],[td["在院"],"在院"],[td["退院"],"退院"],[td["死亡"],"死亡"],[td["日計"],"日計"]]
  temp_list.each do |ca|
    (td["先頭行"]+1..td["最終行"]-1).each do |row|
      write_database("kangodo",nendo,year,month,byoto[sh.cells(row,td["新入院列"]-1).value],ca[1].toutf8,sh.cells(row,ca[0]).value, nil,0,0)
    end
  end

  #看護度
  
  #td=d["看護度"]
  #(td["左上"]["列"]+1..td["右下"]["列"]-1).each do |col|
  #  (td["左上"]["行"]+1..td["右下"]["行"]-1).each do |row|
  #    write_database("kangodo",nendo,year,month,byoto[sh.cells(row,td["左上"]["列"]).value],sh.cells(td["左上"]["行"],col).value.toutf8,sh.cells(row,col).value, nil,0,0)
  #  end
  #	end
  
  #外科合算
  sql="insert into data (nendo, year, month, ka_id, cont, value, memo, nyugai, view) select nendo, year, month, 300, cont, value, memo, nyugai, view from data where ka_id in (3,29) and nendo=? and month=?"
  @db.execute(sql,nendo,month)
  puts "完了"
end #transaction
rescue =>e
  p e
  p e.backtrace
end
