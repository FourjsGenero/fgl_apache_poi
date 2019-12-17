IMPORT JAVA org.apache.poi.ss.formula.functions.Irr         #https://poi.apache.org/apidocs/dev/org/apache/poi/ss/formula/functions/Irr.html
IMPORT JAVA org.apache.poi.ss.formula.functions.FinanceLib  #https://poi.apache.org/apidocs/dev/org/apache/poi/ss/formula/functions/FinanceLib.html

PUBLIC TYPE list_of_values_type DYNAMIC ARRAY OF FLOAT
TYPE java_list_of_values_type ARRAY [] OF FLOAT



FUNCTION irr(list_of_values list_of_values_type, guess FLOAT) RETURNS FLOAT
DEFINE java_list_of_values java_list_of_values_type
DEFINE i INTEGER

    LET guess = nvl(guess, 0.10)

    LET java_list_of_values =  java_list_of_values_type.create(list_of_values.getLength())
    FOR i = 1 TO list_of_values.getLength()
        LET java_list_of_values[i] = list_of_values[i]
    END FOR
    RETURN Irr.irr(java_list_of_values, guess)
END FUNCTION



FUNCTION future_value(rate FLOAT, number_of_periods FLOAT, payment FLOAT, present_value FLOAT, beginning_or_end BOOLEAN) RETURNS FLOAT
    RETURN FinanceLib.fv(rate, number_of_periods, payment, present_value, beginning_or_end)
END FUNCTION

FUNCTION present_value(rate FLOAT, number_of_periods FLOAT, payment FLOAT, future_value FLOAT, beginning_or_end BOOLEAN) RETURNS FLOAT
    RETURN FinanceLib.pv(rate, number_of_periods, payment, future_value, beginning_or_end)
END FUNCTION

FUNCTION number_of_periods(rate FLOAT, payment FLOAT, present_value FLOAT, future_value FLOAT, beginning_or_end BOOLEAN) RETURNS FLOAT
    RETURN FinanceLib.nper(rate, payment, present_value, future_value, beginning_or_end)
END FUNCTION

FUNCTION payment(rate FLOAT, number_of_periods FLOAT, present_value FLOAT, future_value FLOAT, beginning_or_end BOOLEAN) RETURNS FLOAT
    RETURN FinanceLib.pmt(rate, number_of_periods, present_value, future_value, beginning_or_end)
END FUNCTION

FUNCTION npv(rate FLOAT, list_of_values list_of_values_type) RETURNS FLOAT
DEFINE java_list_of_values java_list_of_values_type
DEFINE i INTEGER    

    LET java_list_of_values =  java_list_of_values_type.create(list_of_values.getLength())
    FOR i = 1 TO list_of_values.getLength()
        LET java_list_of_values[i] = list_of_values[i]
    END FOR
    RETURN FinanceLib.npv(rate, java_list_of_values)
END FUNCTION