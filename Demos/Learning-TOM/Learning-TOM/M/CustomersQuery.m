let
    Source = Sql.Database("devcamp.database.windows.net", "WingtipSalesDB"),
    dbo_Customers = Source{[Schema="dbo",Item="Customers"]}[Data],
    RemovedOtherColumns = Table.SelectColumns(dbo_Customers,{"CustomerId", "FirstName", "LastName", "City", "State", "Zipcode", "Gender", "BirthDate", "FirstPurchaseDate", "LastPurchaseDate"}),
    MergedColumns = Table.CombineColumns(RemovedOtherColumns,{"FirstName", "LastName"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Customer"),
    ReplacedFemaleValues = Table.ReplaceValue(MergedColumns,"F","Female",Replacer.ReplaceValue,{"Gender"}),
    ReplacedMaleValues = Table.ReplaceValue(ReplacedFemaleValues,"M","Male",Replacer.ReplaceValue,{"Gender"}),
    ChangedType = Table.TransformColumnTypes(ReplacedMaleValues,{{"FirstPurchaseDate", type date}, {"LastPurchaseDate", type date}, {"BirthDate", type date}}),
    AddedConditionalColumn = Table.AddColumn(ChangedType, "Customer Type", each if [FirstPurchaseDate] = [LastPurchaseDate] then "One-time Customer" else "Repeat Customer"),
    RemovedColumns = Table.RemoveColumns(AddedConditionalColumn,{"FirstPurchaseDate", "LastPurchaseDate"}),
    ChangedType1 = Table.TransformColumnTypes(RemovedColumns,{{"Customer Type", type text}}),
    RenamedColumns = Table.RenameColumns(ChangedType1,{{"City", "City Name"}}),
    AddedCustom = Table.AddColumn(RenamedColumns, "City", each [City Name] & ", " + [State]),
    ChangedType2 = Table.TransformColumnTypes(AddedCustom,{{"City", type text}})
in
    ChangedType2