from fractions import Fraction

from xlfunctions.xlfdate import serial_to_date


base_fmtid = {item[0]:item for item in
    [
    ( 0, ['General'], ['{:g}']),
    ( 1, ['0'], ['{:.0f}']),
    ( 2, ['0.00'], ['{:.2f}']),
    ( 3, ['#,##0'], ['{:,.0f}']),
    ( 4, ['#,##0.00'], ['{:,.2f}']),
    ( 9, ['0%'], ['{:.0%}']),
    (10, ['0.00%'], ['{:.2%}']),
    (11, ['0.00E+00'], ['{:.2E}']),
    (12, ['# ?/?'], ['{} {}/{}']),
    (13, ['# ??/??'], ['{} {}/{}']),
    (14, ['mm-dd-yy'], ['{:%m-%d-%y}']),
    (15, ['d-mmm-yy'], ['{:%d-%b-%y}']),
    (16, ['d-mmm'], ['{:%d-%b}']),
    (17, ['mmm-yy'], ['{:%b-%y}']),
    (18, ['h:mm AM/PM'], ['{:%I:%M %p}']),
    (19, ['h:mm:ss AM/PM'], ['{:%I:%M:%S %p}']),
    (20, ['h:mm'], ['{:%H:%M}']),
    (21, ['h:mm:ss'], ['{:%H:%M:%S}']),
    (22, ['m/d/yy h:mm'], ['{:%m/%d/%y %H:%M}']),
    (37, ['#,##0 ', '(#,##0)'], ['{:,.0f} ', '({:,.0f})']),
    (38, ['#,##0 ', '[Red](#,##0)'], ['{:,.0f} ', '[Red]({:,.0f})']),
    (39, ['#,##0.00', '(#,##0.00)'], ['{:,.2f}', '({:,.2f})']),
    (40, ['#,##0.00', '[Red](#,##0.00)'], ['{:,.2f}', '[Red]({:,.2f})']),
    (45, ['mm:ss'], ['{:%M:%S}']),
    (46, ['[h]:mm:ss'], ['{:%H:%M:%S}']),
    (47, ['mmss.0'], ['{:%M%S.%f}']),
    (48, ['##0.0E+0'], ['{:.1E}']),
    (49, ['@'], ['{}'])
]}

dates_fmt = [*range(14, 23), 45, 46, 47]
fractions_fmt = [12, 13]



def float_to_mixed_fraction(number: float, ndigits: int = 2) -> tuple[int, int, int]:
    """
    Converts a floating-point number into a mixed fraction string representation.

    For example, a number like 12.475 will be returned as "12 19/40".

    Args:
        number: The float or integer to convert.

    Returns:
        A string representing the number as a mixed fraction.
    """
    # # --- Example Usage ---
    # # Your requested example
    # print(f"12.475  ->  '{float_to_mixed_fraction(12.475)}'")

    # # Other examples
    # print(f"5.5     ->  '{float_to_mixed_fraction(5.5)}'")
    # print(f"-3.25   ->  '{float_to_mixed_fraction(-3.25)}'")
    # print(f"0.33333 ->  '{float_to_mixed_fraction(0.33333)}'")
    # print(f"100     ->  '{float_to_mixed_fraction(100)}'")
    # print(f"-0.75   ->  '{float_to_mixed_fraction(-0.75)}'")

    if not isinstance(number, (int, float)):
        raise TypeError("The input must be a number.")

    if number == 0:
        return (0, None, None)

    sign = int(number // abs(number))
    number = abs(number)

    integer_part = int(number)
    decimal_part = number - integer_part

    if decimal_part == 0:
        return sign * integer_part, None, None


    # Convert the decimal part to a fraction, limiting the denominator
    # for a cleaner representation of repeating decimals.
    fraction = Fraction(decimal_part).limit_denominator(10 ** ndigits)

    numerator = fraction.numerator
    denominator = fraction.denominator

    if integer_part == 0:
        # It's just a proper fraction
        return None, sign * numerator, denominator
    else:
        # It's a mixed fraction
        return sign * integer_part, numerator, denominator



def num_fmt(fmt_id:int, value:float | str) -> str:
    fmts = base_fmtid[fmt_id][2]
    if len(fmts) == 1:
        fmts = 3 * fmts
    elif len(fmts) == 2:
        fmts = [*fmts, fmts[0]]
    if isinstance(value, str) or fmt_id == 49:
        ndx = 3
        value = str(value)
    else:
        ndx = 0 if value > 0 else 1 if value < 0 else 2
    try:
        fmt_str = fmts[ndx]
    except IndexError:
        fmt_str = None
    fmt_str = fmt_str or '{}'
    if fmt_id in fractions_fmt:
        value = float_to_mixed_fraction(value, ndigits=fmt_id - 11)
        answ = fmt_str.format(*value)
        return answ.replace('None', '').strip(' /')
    if fmt_id in dates_fmt:
        value = serial_to_date(value)
    return fmt_str.format(value)


