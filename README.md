# Initialisation

// Put require | file here

# Creating a new instance of cJSON

```vba
Sub createInstance()
Dim json As New cJSON

  ' Code Here

Set json = Nothing
End Sub
```

# Examples

Accessing an object

```vba
Dim json As New cJSON

With json
  .create " { 'store': [ { 'id':1, brand: 'Ceega', item: 'T-Shirt' }, { 'id':2, brand: 'Amadeus', item: 'Trousers' } ] } "
  .read

  Debug.Print .rep(.toString("['store'][0]")) ' -> {id: 1, brand: Ceega, item: T-Shirt}
  Debug.Print .rep(.toString("['store'][1]['id']")) ' -> 2
End With

Set json = Nothing
```

Looping through an array

```vba
Dim json As New cJSON
Dim i As Long

With json
  .create " [ { name: 'France', capital: 'Paris' } , { name: 'Spain', capital: 'Madrid' }, { name: 'Italy', capital: 'Rome' } ] "
  .read

  For i = 0 To .map
    Debug.Print .rep(.index(i, "['capital']"))
      ' -> Paris
      ' -> Madrid
      ' -> Rome
  Next
End With

Set json = Nothing
```

Load JSON from a file - test.txt

```json
[
	{
		"name" : "John",
		"age" : 34
	},
	{
		"name" : "Eva",
		"age" : 23
	},
	{
		"name" : "Omar",
		"age" : 43
	},
	{
		"name" : "Rebecca",
		"age" : 30
	},
	{
		"name" : "Miriam",
		"age" : 28
	}
]
```

```vba
Dim json As New cJSON
Dim fileName As String

fileName = "...\test.txt" ' the file above

With json
  .load fileName
  .read

  Debug.Print .rep(.index(0)) ' -> {name: John, age: 34}
    
End With

Set json = Nothing
```

Save JSON to a file - test.txt

```vba
Dim json As New cJSON
Dim fileName As String

fileName = "...\test.txt"

With json
  .create " { 'top': { 'mid': 'a', 'in': [1,2,3], 'go': 'forward' } } "
  .read
  .save fileName
End With

Set json = Nothing
```

Result:

```json
{"top": {"mid": "a", "in": [1, 2, 3], "go": "forward"}}
```

Create object using data from Access db

tbl_data from a database called MyDatabase:

| id        | key           | value  |
| ------------- |:-------------:| -----:|
| 1      | a | a|
| 2      | b      |  real |
| 3      | c      |    test |

```vba
Dim json As New cJSON
Dim fileName As String, sql As String
Dim i As Long

fileName = "C:\Users\nbkn0u6\Desktop\MyDatabase.accdb"
sql = "select * from tbl_data"

With json
  .create "sql", fileName, sql
  .read

  For i = 0 To .map
      Debug.Print .rep(.index(i, "['key']"))
        ' -> a
        ' -> b
        ' -> c
  Next
    
End With

Set json = Nothing
```
