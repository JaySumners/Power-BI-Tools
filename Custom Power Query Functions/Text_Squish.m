let
    fn = (str as text) =>
        if( Text.Contains(str, "  ") )
        then @Text_Squish( Text.Replace(str, "  ", " ") )
        else str,

    fnType = type function
        (
            str as 
                (
                    type text meta
                        [
                            Documentation.FieldCaption = "String",
                            Documentation.FieldDescription = "String or Text"
                        ]
                )
        ) as list meta
            [
                Documentation.Name = "Text.Squish",
                Documentation.LongDescription = "A recurisve replace of repeated spaces. Author: Jay Sumners. Repo: https://github.com/JaySumners/Power-BI-Tools",
                Documentation.Examples =
                {
                    [
                        Description = "Example Description",
                        Code = "Text_Squish(""Some   String"")",
                        Result = "Some String"
                    ]
                }
            ]
in
    Value.ReplaceType(fn, fnType)
