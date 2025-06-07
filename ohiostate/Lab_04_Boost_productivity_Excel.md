
# Lab 4 - Make data-driven R&D decisions with Copilot in Excel

Imagine you're a clinical research lead working on the development of ALZ-4018, a promising therapeutic for Alzheimer's disease. Your responsibility is to analyze clinical trial data and identify patterns or insights that inform the drug development strategy. In this hands-on lab, you'll use Copilot in Excel to explore and analyze metrics from ongoing studies.

You'll start by getting an overview of the dataset and identifying key metrics. Next, you'll analyze participant trends, cognitive performance data, and biomarker response. Additionally, you'll examine relationships between dropout rates and adverse events, and generate a summary of clinical findings to share with stakeholders.

## Sample file

Throughout this Lab, we'll craft prompts for Microsoft 365 Copilot that reference the file **ALZ4018_Clinical_Study_Data.xls**. It's an Excel spreadsheet that contains sample data representing results of a ficticious clinical trial.

## Exercise 1 - Explore the dataset

You must first understand the overall trends in your clinical trial. Start by getting a summary and identifying the metrics that can guide your R&D decisions.

1. Open the sample file in Excel from your OneDrive.
2. Select the **Copilot** icon from the **Home** ribbon to open the Copilot pane.
3. Enter the prompt:  
   `Summarize the dataset and provide an overview of key clinical trial metrics.`
4. Copilot will return a summary and optionally create a table. Select **Add to a new sheet**.
5. Review the summary, then return to Sheet 1.

## Exercise 2 - Identify participant trends

Tracking participant enrollment and dropout is key for trial planning.

1. In the Copilot pane, enter the prompt:  
   `Show a line chart of Participants Enrolled over the months.`
2. Review the chart and optionally add it to a new sheet.
3. Next, return to Sheet1 and enter:  
   `Highlight the three months with the highest number of adverse events reported.`
4. Apply the formatting.

## Exercise 3 - Analyze cognitive outcomes

Understanding therapeutic impact on cognitive decline is critical.

1. Prompt Copilot with:  
   `Create a line chart of Cognitive Score Change across the year.`
2. Review and optionally add the chart to a new sheet.
3. To isolate trends for a given quarter, enter:  
   `Summarize the average Cognitive Score Change for June, July, and August.`

## Exercise 4 - Add dropout rate

Identify the dropout rate

1. Prompt Copilot with:  
   `Add a column named "Monthly Change in Patients" that shows the difference in "Patients Enrolled" compared to the previous month. Use N/A for January.'
2. Review the extra column and **Insert column** if appropriate.

## Exercise 5 - Investigate correlation between dropout rate and adverse events

Ensure that safety issues are well understood and tracked.

1. Prompt Copilot with:  
   `Show any correlation between Adverse Events Reported and Monthly Change in Patients.`
2. Review the chart and summary, then select **Add to sheet** if appropriate.

## Exercise 6 - Generate key insights

Summarize the findings for internal reports or regulatory reviews.

1. Enter the following prompt:  
   `Provide a summary of key insights from the ALZ-4018 clinical trial dataset.`

## Exercise 7 - Share findings with the team

Use Copilot in Outlook to communicate insights with collaborators and sponsors.

1. Copy the Copilot-generated summary from Excel.
2. Open Microsoft Outlook and start a new email.
3. Paste the summary into the body of the email.
4. Select the Copilot icon and enter the prompt:  
   `Draft a professional email to my team summarizing the key findings from the ALZ-4018 clinical trial data analysis.`
5. Review the draft and select **Replace** if appropriate.

This lab has helped you use Microsoft 365 Copilot in Excel to efficiently analyze clinical data, identify research patterns, and communicate findings. Continue experimenting with prompts for deeper insights as your trial progresses.

**End of Lab 4**
