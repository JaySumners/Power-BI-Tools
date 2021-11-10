let
    Source = 
        Table.TransformColumnTypes(
            Table.FromColumns(
                {
                    {1..8}
                },
                {"myValue"}
            ),
            {{"myValue", Int64.Type}}
        ),
    Rcd = 
        [
            #"[myValue]<2" = "Test1",
            #"[myValue]<3" = "Test1.5",
            #"[myValue]>5" = "Test2"
        ],
    NewCol = Table.AddColumn(
        Source,
        "Test",
        each Switch(_, Rcd, "default")
    )
in
    NewCol
