let
    Source = Sql.Database("devcamp.database.windows.net", "WingtipSalesDB"),
    dbo_InvoiceDetails = Source{[Schema="dbo",Item="InvoiceDetails"]}[Data],
    ExpandedInvoices = Table.ExpandRecordColumn(dbo_InvoiceDetails, "Invoices", {"InvoiceDate", "CustomerId"}, {"InvoiceDate", "CustomerId"}),
    RemovedColumns = Table.RemoveColumns(ExpandedInvoices,{"Products"}),
    ChangedType = Table.TransformColumnTypes(RemovedColumns,{{"InvoiceDate", type date}, {"SalesAmount", Currency.Type}})
in
    ChangedType