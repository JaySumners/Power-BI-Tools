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
            #"[myValue]>5" = "Test2"
        ],
    NewCol = Table.AddColumn(
        Source,
        "Test",
        each 
            let
                pos = 
                    List.PositionOf(
                        List.Transform(
                            Record.FieldNames(Rcd), 
                            (exp) => 
                                Expression.Evaluate(
                                    Expression.Identifier("x") & exp,
                                    [x = _]
                                )
                        ),
                        true,
                        1
                    ),
                ret = 
                    if pos < 0 then "default" else Record.FieldValues(Rcd){pos}
            in
                ret
    )
in
    NewCol
