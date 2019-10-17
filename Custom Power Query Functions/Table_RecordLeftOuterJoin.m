let
    fn = (tbl as table, rcd as record, join_column as text, new_col_name as text) =>
        if 
            List.Contains(
                Table.ColumnNames(tbl),
                new_col_name
            )
        then 
            error Error.Record(
                "Invalid Parameter",
                "Duplicate Column Name", 
                new_col_name & " (the new column name) already exists in the table."
            )
        else
            let
                type_rcd = Value.Type(Record.ToList(rcd){0}),

                Temp_Rename =
                    Table.RenameColumns(
                        tbl,
                        {{join_column, "join"}}
                    ),

                Source = 
                    Table.AddColumn(
                        Temp_Rename,
                        new_col_name,
                        each
                            if [join] = null
                            then null
                            else
                                Record.FieldOrDefault(
                                    rcd,
                                    [join],
                                    null
                                ),
                        type_rcd  
                    ),

                Undo_Rename =
                    Table.RenameColumns(
                        Source,
                        {{"join", join_column}}
                    )      
            in
                Undo_Rename,

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
                    join_column as
                        (
                            type text meta 
                                [
                                    Documentation.FieldCaption = "Join Column",
                                    Documentation.FieldDescription = "No Description"
                                ]
                        ), 
                    new_col_name as
                        (
                            type text meta 
                                [
                                    Documentation.FieldCaption = "New Column Name",
                                    Documentation.FieldDescription = "No Description"
                                ]
                        )
                ) as list meta
                    [
                        Documentation.Name = "Table_RecordLeftOuterJoin",
                        Documentation.LongDescription = "A left outer join using a record. This process is much faster for a one-to-many or many-to-one join than using the standard ""Merge"". Author: Jay Sumners. Repo: https://github.com/JaySumners/Power-BI-Tools",
                        Documentation.Examples =
                        {
                            [
                                Description = "Example Description",
                                Code = "Table_RecordLeftOuterJoin(mytable,myrecord,""Column1"", ""NewColumn"")",
                                Result = "a table"
                            ]
                        }
                    ]
        in
            Value.ReplaceType(fn, fnType)