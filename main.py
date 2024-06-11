import win32com.client
ol = win32com.client.Dispatch('Outlook.Application')

olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

newmail.Subject = 'Testing Mail'
newmail.To = 'desa2thomas@gmail.com'
HTML = """
<body>
    <p>Hello, <br/><br/> 
        As the Brampton Board of Trade continues to refine our database, we want to ensure that everything is correctly 
        listed within our system so that we can effectively serve you better, allowing you to achieve the greatest possible success.
        <br/><br/> 

        Reconciling our database enables us to:
        <ul> 
            <li>Ensure the current contacts (Reps) are listed and categorized to maximize information sharing</li>
            <li>Ensure physical and mailing addresses are correct and reflected in our online directory</li> 
        </ul>
        </br></br> 
        
        Below, you will find a list containing the current representatives listed within our database. <br/><br/> 
     
"""

#Add logic to iterate through CSV and add the data from every row into this section of the email

HTML += """
    Please confirm your organization's information by <a href= https://www.surveymonkey.com/r/R9M862L> following the survey link <a> 
    and filling out the relevant fields of information.<br/><br/> 

    Completion of the survery will enter members into a draw for a chance to win one of our fantastic prizes:
    <ol>
        <li>Half price BBOT Membership on the next annual renewal.</li>
        <li>Two complimentary tickets to a BBOT Signature Event of their choosing during the 2024-2025 year.</li> 
    </ol>
    Thank you for your time, <br/><br/> e

    <strong>Hana Brissenden</strong> <br/>
    Brampton Board of Trade <br/>
    36 Queen St E #101, Brampton, ON L6V 1A2 <br/>
    T: (905) 451-1122
    <a href = www.bramptonbot.com> </a> 
"""
HTML += """</p></body>"""


newmail.HTMLBody = HTML
newmail.Display()
