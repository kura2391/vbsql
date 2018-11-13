# VbSql

vbsqlは、visual basicでのsqlserver接続とsql文実行を手助けします

## 注意事項

- 現在v0.5.0
- sqlのすべての機能を利用できるわけではありません。ご了承ください。

## インストール方法

その1
* このプロジェクトをダウンロードします
* visual studioでvbsql.slnを起動し、ビルドを実行します
* ビルド後、作成された.dllファイルを、インストールしたいプロジェクトの参照に追加します。

その2
* インストールしたいプロジェクトの参照に、bin/release内のVbsql.dllファイルを追加します。

## 使い方

### 全体の説明
- Select, Insert, Update, Deleteはすべてつなげて実行できます。具体例としては、Vbsql.Deleteをご覧ください。
- Connectionクラスを作成し、各SQLクラスに渡すことで実行できます。

### Vbsql.Connection

~~~vb
'2通りあります
'connectionString ... sqlserver接続文字列

Dim conn as new Vbsql.Connection(connectionString)

Dim conn As New Vbsql.Connection("server", "userid", "password", "initialCatalog")
~~~

### Vbsql.Select
```vb
'conn ... VbSql.Connection クラス
Dim select As New vbsql.Select(conn)

' SELECT * FROM test 
' INNER JOIN test2 ON test.number = test2.number 
' WHERE id = 500 and date > "2018-10-9" 
' ORDER BY date asc;
select.from("test")
select.select({"*"}) 
select.join("test2", "test.number = test2.number",Vbsql.Jointype.INNER)
select.where("id = ? and date > ?",{"500","2018-10-9"}) 
select.orderBy("date asc") 

'結果がDataTableとして返されます
Dim dt as DataTable = sel.execute() 

' 特記事項
' select ... 欲しい列名を配列で入力します。例) {"number","date"}
' join ... 第三引数にjoinのtypeを指定します。例)Vbsql.Jointype.LEFT
' where ... ユーザー入力欄は ? とし、第二引数に配列として入力します。
```

### Vbsql.Insert
例1
```vb

Dim insert As New Vbsql.Insert(connection)

' INSERT INTO test(date,text) 
' VALUES("2018-10-15","変更");SELECT SCOPE_IDENTITY();
Dim ht As New Hashtable
ht("date") = "2018-10-15"
ht("text") = "変更"
insert.into("test")
insert.value(ht)
insert.lastInsertId() 

' lastInsertIdを有効にしていた場合,lastInsertIdを返します。
dim lastinsertid as Integer = insert.execute()

' 特記事項
' 列名とその値の組はHashtableで渡します。
' NULLを挿入したい場合、DBNull.Valueを入れてください。
' 例) ht("text") = DBNull.Value
' lastInsertId ... 「SELECT SCOPE_IDENTITY();」を入れる際はこの関数を記入します。
' lastInsertIdの意味については、インターネットで検索してください。
```

例2(複数行のINSERT)
```vb

' INSERT INTO test(text,number)
' VALUES
' ("0のてすと","10000"),
' ("1のてすと","10001"),
' ("2のてすと","10002")
Dim insert As New Vbsql.Insert(connection)
insert.into("test")
Dim dt As New DataTable
dt.Columns.Add("text")
dt.Columns.Add("number")
For i As Integer = 0 To 2
    Dim row As DataRow = dt.NewRow
    row("text") = i & "のてすと"
    row("number") = i + 10000
    dt.Rows.Add(row)
Next

insert.values(dt)
insert.execute()

' 特記事項
' 複数行のINSERTを行いたい場合、DataTableにそのデータを入れます。
' NULLの挿入もできます。
' 例) row("number") = DBNull.Value
```

※Insert.Value(ht as Hashtable)とInsert.Values(dt as Datatable)は同時には使用できません。
最後に実行されたValue関数( or Values関数)が適用されます。

### Vbsql.Update
```vb

' UPDATE test 
' SET text="変更します", number=9923 
' WHERE id = 1003 and date = 2018-10-22
Dim update As New Vbsql.Update(connection)
Dim ht As New Hashtable
ht("text") = "変更します"
ht("number") = 9923
update.table("test")
update.set(ht)
update.where("id = ? AND date = ?", {"1003", "2018-10-22"})
update.execute()

' 特記事項
' 変更内容をHashtableで渡します。
' DBNull.Valueを入力することで、NULLへの変更も可能です。
' 例)ht("text") = DBNull.Value
' where ... ユーザー入力欄は ? とし、第二引数に配列として入力します。

```


### Vbsql.Delete
```vb
Dim delete As New vbsql.Delete(connection)

' DELETE FROM test WHERE id=1003
delete.from("test").where("id = ?", {"1003"}).execute() 

'特記事項
' where ... ユーザー入力欄は ? とし、第二引数に配列として入力します。
```

