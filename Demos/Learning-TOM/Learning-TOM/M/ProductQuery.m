let
    Source = Sql.Database("devcamp.database.windows.net", "WingtipSalesDB"),
    dbo_Products = Source{[Schema="dbo",Item="Products"]}[Data],
    RemovedOtherColumns = Table.SelectColumns(dbo_Products,{"ProductId", "Title", "Description", "ProductCategory", "ProductImageUrl"}),
    RenamedColumns = Table.RenameColumns(RemovedOtherColumns,{{"Title", "Product"}}),
    SplitColumnByDelimiter = Table.SplitColumn(RenamedColumns, "ProductCategory", Splitter.SplitTextByDelimiter(" > ", QuoteStyle.Csv), {"ProductCategory.1", "ProductCategory.2"}),
    ChangedType = Table.TransformColumnTypes(SplitColumnByDelimiter,{{"ProductCategory.1", type text}, {"ProductCategory.2", type text}}),
    RenamedColumns1 = Table.RenameColumns(ChangedType,{{"ProductCategory.1", "Category"}, {"ProductCategory.2", "Subcategory"}})
in
    RenamedColumns1