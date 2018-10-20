# VbSql

vbsqlは、visual basicでのsqlserver接続とsql文実行を手助けします

## インストール方法

* このプロジェクトをダウンロードします
* visual studioでvbsql.slnを起動し、ビルドを実行します
* ビルド後、作成された.dllファイルを、インストールしたいプロジェクトの参照に追加します。

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
select.from("test") ' table名
select.select({"*"}) ' 列名をString()で指定　例 {"id","date","text"}
select.where("id = ? and date > ?",{"500","2018-10-9"}) ' 変数を ? で、第二引数にその値をString()で。
select.orderBy("date asc") 'Order By句をstringで
Dim dt as DataTable = sel.execute() '結果がDataTableとして返されます
```

VbSql.Update
```vb
Dim update As New vbsql.Update(connection)

Dim param As New Dictionary(Of String, String) 'setする値をDictionary で保存します
param.Add("text", "変更しましたere") ' text="変更しましたere"
param.Add("number", "9923") ' number=9923 

update.table("test") ' table名の指定
update.set(param) ' 先ほどのparameterをセット
update.where("id = ? AND date = ?", {"1003", "2018-10-15"}) '条件を記入
update.execute() 
```

VbSql.Insert
```vb

Dim ins As New Vbsql.Insert(connection)

Dim param As new Dictionary(Of String,String) 
param.Add("date","2018-10-15")

ins.into("test") 'table名の指定
ins.values(param) ' 上述のパラメータを設定
ins.lastInsertId() ' これにより、返り値がLastInsertidの値(select SCOPE_IDENTITY()の値)になります
dim lastinsertid as Integer = ins.execute()

```

VbSql.Delete
```vb
Dim delete As New vbsql.Delete(connection)

'delete,insert,select,update文はこのようにくっつけて実行もできます
delete.from("test").where("id = ?", {"1003"}).execute() 

```
