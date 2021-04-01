import pandas as pd
import numpy as np
import pdfkit

VENUE = 'warwick'

DAY = '30'
MONTH = '03'
YEAR = '2021'
HOUR = '16'
MIN = '45'

FILE = f'WAR_{DAY}{MONTH}_{HOUR}{MIN}.xlsx'

TITLE = "2m 4f 0y Air Wedding Open Hunters' Chase"
DATE = f'{DAY}th March 2021'
TIME = f'{HOUR}:{MIN}'

dataframe = pd.read_excel(FILE)
rounded_df = dataframe.round({'Last 3f (%)': 2})
new_df = rounded_df.replace(np.nan, '', regex=True)


html_table = new_df.to_html(index=False, classes="horses", border=0)
# print(html_table)

html_string = '''
<!DOCTYPE html>
<html>
<head>
<title>Post Race Reports</title>
<style>

</style>
<link type="text/css" rel="stylesheet" href="style.css">
</head>
<body>
    <div class="all">
        <div class="header1">
            <div class="race-logo"><img src="logos/{venue}.png" alt="Kempton Logo" style="width:180px;height:80px;">
            </div>
            <div class="title-date-container">
                <h2>{title}</h2>
                <h3>{date}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Time: {time}</h3>
            </div>
            <div class="racingTV-logo"><img src="racingTV.png" alt="Racing TV logo"
                style="width:300px;height:81px;">
            </div>
        </div>

     {table}

    </div>
    <div class="lower-banner">
        <div class="times-accuracy">
            <p>Times accurate to +/- 0.2s or more excluding the location accuracy +/-0.5m (approximately 0.03s).
                </br>One horse length is run in approximately 0.167 seconds on good or firmer ground, 0.18 seconds on
                good to soft ground and 0.2 seconds or more on soft.</p>
        </div>
        <div class="email" style="text-align:right;"><a href="mailto:sectionals@racingtv.com">For enquiries, please contact sectionals@racingtv.com</a>
        </div>
    </div>
</div>
</body>
</html>
'''
html_pdf_string = '''
<!DOCTYPE html>
<html>
<head>
<title>Post Race Reports</title>

<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400&display=swap');

        html,
        body {{margin: 0px;
            padding: 0px;
            font-family: 'Source Sans Pro', sans-serif;}}
            
        .header {{}}
        
        .title-date-container {{color: rgb(38, 33, 97);}}

        .times-accuracy {{color: grey;
            margin-left: 0;}}

        .email {{display: flex;
            justify-content: flex-end;
            font-size: 17px;
            right: 0;}}

        a:link {{color: rgb(38, 33, 97);
            text-decoration: none;}}

        a:hover {{text-decoration: underline;
            color: rgb(38, 33, 97);}}

        .horses {{font-family: 'Source Sans Pro', sans-serif;
            border-collapse: collapse;
            width: 100%;}}

        .horses td,
        .horses th {{border: 1px solid #ddd;
            padding: 8px;
            font-family: 'Source Sans Pro', sans-serif;}}

        .horses tr:nth-child(even) {{background-color: #f2f2f2;}}

        .horses tr:hover {{background-color: #ddd;}}

        .horses th {{padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #58575c;
            color: white;}}


        td.runnername {{column-width: 200px;}}
        td.statcat {{display: none;}}

        th.statcat {{display: none;}}
        th.runnernumber {{display: none;}}
        td.runnernumber {{display: none;}}
</style>

<meta name="pdfkit-orientation" content="Landscape"/>

</head>
<body>
    <div class="all">
        <div class="header">
            <div class="race-logo" style="float:left;width:15%;"><img src="/Users/dimitrivanloey/PycharmProjects/from_excel_to_html_jumps/kempton.png" 
            alt="Kempton Logo" style="width:180px;height:90px;">
            </div>
            <div class="title-date-container" style="float:left;width:60%;margin-top:-20px;text-align:center;">
                <h1 style="display:inline-block;width: 750px;height: 15px;text-align:left;">{title}</h1>
                <h3 style="display:inline-block;width: 750px;height: 15px;text-align:left;">{date}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Time: {time}</h3>
            </div>
            
            <div class="racingTV-logo" style="float:right;width:23%;">
            <img src="/Users/dimitrivanloey/PycharmProjects/from_excel_to_html_jumps/racingTV.png" alt="Racing TV logo" 
                style="width:300px;height:85px;">
            </div>
        </div>

     {table}

    </div>
    <div class="lower-banner">
        <div class="times-accuracy">
            <p>Times accurate to +/- 0.2s or more excluding the location accuracy +/-0.5m (approximately 0.03s).
                </br>One horse length is run in approximately 0.167 seconds on good or firmer ground, 0.18 seconds on
                good to soft ground and 0.2 seconds or more on soft.</p>
        </div>
        <div class="email" style="text-align:right;"><a href="mailto:sectionals@racingtv.com">sectionals@racinguk.tv</a>
        </div>
    </div>
</div>
</body>
</html>
'''

with open('website.html', 'w') as website_file:
    website_file.write(html_string.format(table=new_df.to_html(index=False, classes="horses", border=0), title=TITLE,
                                          date=DATE, time=TIME, venue=VENUE))

with open('pdf_website.html', 'w') as pdf_website_file:
    pdf_website_file.write(html_pdf_string.format(table=new_df.to_html(index=False, classes="horses", border=0),
                                                  title=TITLE, date=DATE, time=TIME))

formatted_html_pdf_string = html_pdf_string.format(table=new_df.to_html(index=False, classes="horses", border=0),
                                                   title=TITLE, date=DATE, time=TIME)

options = {
  "enable-local-file-access": None
}

# To PDF
# css = ["style.css"]
# pdfkit.from_file('pdf_website.html', 'kempton-park_20210171_1850.pdf', options=options, css=css)
# pdfkit.from_string(formatted_html_pdf_string, f'kempton-park_{YEAR}{MONTH}{DAY}_{HOUR}{MIN}.pdf', options=options)