'''
- Frequency: Monthly / Quarterly
- 3 levels of reporting: Global, Regions (US/EMEA/APAC), individual IM
- # project signed for the period (added to monday)
- # project started for the period
- # project completed for the period
- # canceled project for the period
- # paused/on hold project for the period
- Type of tech setup: connectors (target <3 months) or custom (target <6 months)
- avg length between date project added to Monday and Kick-off date
- avg length of the start date and due date of projects for the period
- avg length of the completed projects for the period
'''


'''
First Stage:
    Monday
        1st Step:
        - Gather project data
        2nd Step:
        - Group project data by region
        3rd Step:
        - Use grouped projects by region to then group by frequency (quarterly/monthly)
        4th Step:
        - Break down our grouped by frequency project data to gather numbers for KPIs
            Quarterly:
            Q1:
            --'projects_signed': 40,
            --'projects_started': 25,
            --'projects_completed': 70,
            --'canceled_projects': 5,
            --'paused_projects': 2,

            Monthly:
            January:
            --'projects_signed': 15,
            --'projects_started': 9,
            --'projects_completed': 3,
            --'canceled_projects': 0,
            --'paused_projects': 1,

Second Stage:
    Google
        5th Step:
        - Need to use KPI data to send to Google Slides

'''