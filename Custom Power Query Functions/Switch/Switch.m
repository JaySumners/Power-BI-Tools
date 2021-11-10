(
    eachRecord as record,
    conditionsRecord as record,
    optional default as any
) =>
let
    position = 
        List.PositionOf(
            List.Transform(
                Record.FieldNames(conditionsRecord), 
                (exp) => 
                    Expression.Evaluate(
                        Expression.Identifier("x") & exp,
                        [x = eachRecord]
                    )
            ),
            true,
            1
        ),
    ret = 
        if position < 0 then default else Record.FieldValues(conditionsRecord){position}
in
    ret
