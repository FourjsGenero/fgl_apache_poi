IMPORT FGL fgl_financial

MAIN
DEFINE list_of_values fgl_financial.list_of_values_type
DEFINE irr FLOAT
DEFINE i INTEGER

    DISPLAY ""
    DISPLAY "*** Simple IRR Tests ***"
    CALL list_of_values.clear()
    LET list_of_values[1] = -100
    LET list_of_values[2] = 110
    LET irr = fgl_financial.irr(list_of_values,NULL)
    DISPLAY SFMT("IRR is %1", irr)

    CALL list_of_values.clear()
    LET list_of_values[1] = -1000000
    LET list_of_values[2] = 550000
    LET list_of_values[3] = 550000
    LET irr = fgl_financial.irr(list_of_values,NULL)
    DISPLAY SFMT("IRR is %1", irr)

    DISPLAY ""
    DISPLAY "*** Financial tests ***"
    # The following are all based around a mortgage of 100,000 at 3.75% for 20 years with monthly payments
    # strictly speaking I think I should not be using .03.75/12 but should be taking compounding into account
    # in which payment will be 593.10

    CALL list_of_values.clear()
    LET list_of_values[1] = -100000
    FOR i = 1 TO 240
        LET list_of_values[i+1] = 592.89
    END FOR
    LET irr = fgl_financial.irr(list_of_values,0.10/12)
    DISPLAY SFMT("IRR is %1, should be close to 3.75%%", irr*12)

    DISPLAY "*** NPV Test ***"
    CALL list_of_values.clear()
    FOR i = 1 TO 240
        LET list_of_values[i] = 592.89
    END FOR
    DISPLAY SFMT("NPV is %1, should be close to 100,000",fgl_financial.npv(0.0375/12, list_of_values))

    DISPLAY "*** Number of period test ***"
    DISPLAY SFMT("Number of periods is %1, should be close to 240",fgl_financial.number_of_periods(.0375/12,-592.89, 100000,0,FALSE))

    DISPLAY "*** Payment Test ***"
    DISPLAY SFMT("Payment is %1, should be close to -592.89", fgl_financial.payment(.0375/12,240,100000,0, FALSE))

    DISPLAY "*** Present Value ***"
    DISPLAY SFMT("Present Value is %1, should be close to 100,000", fgl_financial.present_value(.0375/12,240,-592.89,0, FALSE))

    DISPLAY "*** Future Value ***"
    DISPLAY SFMT("Future value is %1, should be close to 0",fgl_financial.future_value(.0375/12,240,592.89,-100000, FALSE))
END MAIN