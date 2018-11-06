# VbSql

vbsqlは、visual basicでのsqlserver接続とsql文実行を手助けします

## インストール方法

その1
* このプロジェクトをダウンロードします
* visual studioでvbsql.slnを起動し、ビルドを実行します
* ビルド後、作成された.dllファイルを、インストールしたいプロジェクトの参照に追加します。

その2
* インストールしたいプロジェクトの参照に、bin/release内のVbsql.dllファイルを追加します。

## 使い方

VbSql.Connection

~~~vb
'connectionString ... sqlserver接続文字列
Dim conn as new VbSql.Connection(connectionString)
'or あまり変わりませんが...
Dim connection As New Connection("server", "userid", "password", "initialCatalog")
~~~

VbSql.Select
```vb
'conn ... VbSql.Connection クラス
Dim select As New vbsql.Select(conn)

' table名
select.from("test")
' 列名をString()で指定　例 {"id","date","text"}
select.select({"*"}) 
' JOIN test2 ON test.number = test2.number となります。LEFT,RIGHT,FULLJOINについては、第三引数にその名前を入れます。
select.join("test2", "test.number = test2.number",Vbsql.Jointype.INNER)
' 変数を ? で、第二引数にその値をString()で。
select.where("id = ? and date > ?",{"500","2018-10-9"}) 
'Order By句をstringで
select.orderBy("date asc") 

'結果がDataTableとして返されます
Dim dt as DataTable = sel.execute() 
```

VbSql.Update
```vb
Dim update As New vbsql.Update(connection)

'setする値をDictionary で保存します
Dim param As New Dictionary(Of String, String) 
' text="変更しましたere"
param.Add("text", "変更しましたere") 
' number=9923 
param.Add("number", "9923") 

' table名の指定
update.table("test")
' 先ほどのparameterをセット
update.set(param) 
'条件を記入
update.where("id = ? AND date = ?", {"1003", "2018-10-15"}) 

update.execute() 
```

VbSql.Insert
```vb

Dim ins As New Vbsql.Insert(connection)

' update と同様に、Dictionaryクラスに必要なデータを入れていきます。
Dim param As new Dictionary(Of String,String) 
param.Add("date","2018-10-15")

'table名の指定
ins.into("test") 
' 上述のパラメータを設定
ins.values(param) 
' 返り値をLastInsertidの値(select SCOPE_IDENTITY()の値)にします(省略可能)
ins.lastInsertId() 

dim lastinsertid as Integer = ins.execute()

```

VbSql.Delete
```vb
Dim delete As New vbsql.Delete(connection)

'delete,insert,select,update文はこのようにくっつけて実行もできます
delete.from("test").where("id = ?", {"1003"}).execute() 

```
