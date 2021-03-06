let
    fn = (tbl as table, rcd as record, replace_in_column as text) =>
        let
            type_rcd = Value.Type(Record.ToList(rcd){0}),

            Source = 
                Table.TransformColumns(
                    tbl,
                    {
                        replace_in_column,
                        each
                            if _ = null 
                            then null
                            else
                                Record.FieldOrDefault(
                                    rcd,
                                    _,
                                    null
                                ),
                        type_rcd
                    }
                )   
        in
            Source,

    fnType = type function
        (
            tbl as
                (
                    type table meta 
                        [
                            Documentation.FieldCaption = "Table",
                            Documentation.FieldDescription = "No Description"
                        ]
                ), 
            rcd as
                (
                    type record meta 
                        [
                            Documentation.FieldCaption = "Record",
                            Documentation.FieldDescription = "No Description"
                        ]
                ), 
            replace_in_column as
                (
                    type text meta 
                        [
                            Documentation.FieldCaption = "Column to Transform",
                            Documentation.FieldDescription = "No Description"
                        ]
                )
        ) as list meta
            [
                Documentation.Name = "Table_RecordReplaceValues",
                Documentation.LongDescription = "A multi-value replace using an established or manually entered record. NOTE: You cannot replace nulls using a record. Author: Jay Sumners. Repo: https://github.com/JaySumners/Power-BI-Tools",
                Documentation.Examples =
                {
                    [
                        Description = "Example Description",
                        Code = "Table_RecordReplaceValues(mytable, myrecord,""column I want to replace value in"")",
                        Result = "a table"
                    ]
                }
            ]
in
    Value.ReplaceType(fn, fnType)