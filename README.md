# Mortgage-Point-Calculator

Mortgage points are confusing. After I looked into it, however, I was surprised by the sheer magnitude of obfuscation of their value. If this problem had been given to me in Industrial Economics class, I would have gotten it wrong. Figuring out some coherent method of valuing them is actually a reasonably challenging problem.

It's easy to find boiler-plate formulas for dealing with most basic [present, annual, and future](http://www.ajdesigner.com/phppresentworth/present_worth_annual_payment_equation.php) value conversions. Since points lower a mortgage's rate, and thus the payment, it initially appears like a simple calculation. It is not. Firstly, you have the basic difference in rate, which calls for converting annual value to present value. Secondly, the remaining principle of the loan is different from the reference scenario to the scenario of buying points. Finially, the rate of return depends strongly on if and when you leave the mortgage, in which case the extra cash spent up-front is probably wasted. Every one of these factors will be changed depending on what your own personal expected rate of return is from alternative investments too.

For a particular scenario, you can get a sense of typical rates and the prices of points from places on the internet like [lenderfi](www.lenderfi.com). Our goal here is to make sense of what you get for that initial investment in terms of higher closing costs.

## Excel

Functions are implemented in Visual Basic so that they can be used in an Excel spreadsheet saved as a .xlsm file, which is needed to enable macros. Make sure you have added the developer ribbon, start a new file, click on visual basic, in the new window right click on VBAProject and then insert > module and paste the code into there. After that you can use the "=" sign in cells to use the functions with real numbers in your spreadsheet.

