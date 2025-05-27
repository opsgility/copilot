
# Craft effective prompts for Microsoft 365 Copilot

## Lesson 1: Get started with Microsoft 365 Copilot


## Lab 

In this lesson, you learn how to craft effective, contextual prompts for Microsoft 365 Copilot to **summarize** information in Microsoft 365 apps. If you're looking to enhance your prompting skills for Copilot, this lesson equips you with the knowledge and techniques to craft prompts that yield accurate and helpful results from Microsoft 365 Copilot.

The focus of this lesson is **simplifying**, **summarizing**, **understanding**, and **compiling** information using Copilot in Microsoft 365 apps such as **Word**, **PowerPoint**, **Teams**, **Outlook**, and others. You learn how to use these basic summarization capabilities, but also how to craft an efficient prompt that contains all the elements to generate the desired results.


By the end of this lesson, you'll be able to:

1. Identify the key elements of an effective prompt and apply them to your own prompts.

1. Prompt Copilot to summarize or extract key information in Word documents, Excel tables, and PowerPoint presentations.

1. Summarize chats and meetings to look for key action items with Copilot in Teams.

1. Task Copilot in Outlook with summarizing emails to look for action items or mentions.

1. Compile information from multiple documents and generate a combined summary with Microsoft 365 .

Throughout the lesson, you'll see examples of **simple** prompts that are improved with different techniques to be **good** prompts, then **better** prompts, and then the **best** version of those prompts. Dissecting prompts in this way helps you learn about the capabilities of Microsoft 365 Copilot and reinforces the importance of good technique at the same time.

Completing this lesson provides you with skills to **enhance productivity** by effectively communicating expectations to Copilot, saving time and effort. You'll also develop **improved prompting skills** to craft clear and concise prompts, enhancing productivity in Microsoft 365 apps.



## Uploading Files to OneDrive 

Follow the steps below to upload all files needed to **OneDrive**:

1. This course uses sample files throughout the exercises. Download the following .zip file to your computer and extract to a new folder called CopilotDocs on your desktop.

    ```
    https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005/SampleFiles/CopilotDocs.zip
    ```


1. Open **Microsoft Edge** then navigate to the below URL. 

    ```
    https://onedrive.live.com/login
    ```

2. If prompted, login with your work credentials. 

3. In **OneDrive**, in the top-left corner, select **+ Add new** > **Folder upload**.

    [](media/add_new.png)

4. In **File Explorer**, select **Desktop** and open the **CopilotDocs** folder and click Upload again to upload the folder to OneDrive.

5. Leave **Edge** open and move on to the next task.



### Exercise 1: Simplify and extract key information with Copilot in Word

In this exercise, you'll use Microsoft 365 Copilot in Word to summarize and extract insights from a market analysis report. You'll learn how to refine your prompts by adding context, specifying document sections, and setting clear expectations—transforming basic prompts into precise instructions that deliver more focused and useful responses.

1. Once you are signed into Word, click **Blank document**.

1. To start using Microsoft 365 Copilot in Word, you can open the **Copilot** pane by selecting the Copilot icon on the far right of the ribbon's **Home** tab.

    ![Screenshot of the Copilot icon in the Word ribbon.](media/summarize_copilot-ribbon-word.png)

    This helpful feature provides answers to questions—broad or specific—about your document. Have a back-and-forth discussion to iterate and refine your results, get a summary or specific information about the document content, or ask it to generate ideas, tables, or lists that you can copy and insert into your document.


1. In Word, open the **Market Analysis Report for Mystic Spice Premium Chai Tea.docx** document within the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder. 

1. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

     In this simple prompt, you start with the basic **Goal**: _to summarize a Word document._ 
    
    However, there's no information about why the document needs to be summarized or what the summary is needed for.

    Start with a basic prompt by defining your goal. For example, you might say:

    ```
    Summarize this Word document.
    ```

    To make it a good prompt, add context to help Copilot understand the purpose. For instance:

    ```
    Summarize this Word document with a brief overview of the main points to discuss with my team during tomorrow's Sales meeting.
    ```

    To create a better prompt, specify the source or section you'd like summarized. This helps Copilot focus more precisely. For example:

    ```
    Summarize the section on Competitive Analysis with a brief overview for my Sales meeting.
    ```

    For the best prompt, set clear expectations about format or detail level. This guides Copilot’s output. For example:

    ```
    Summarize the section on Competitive Analysis. Please keep the summary to 5 key points and use simple language."
    ```

    Here, we have a better crafted prompt: 

    ```
    Summarize the section on Competitive Analysis in this Word document with a brief overview of the main points to discuss with my team during the tomorrow's Sales meeting. Please keep the summary to 5 key points and use simple language.
    ```

    This prompt has all the details it needs - **Goal**, **Context**, **Source**, and **Expectations** - so Copilot can give you the answer you're looking for.



### Exercise 2: Identify key information and summarize with Copilot in PowerPoint

In this exercise, you'll use Microsoft 365 Copilot in PowerPoint to summarize specific sections of a presentation. You'll learn how to craft prompts that guide Copilot to focus on the right slides, apply professional formatting, and tailor the summary for a business audience—enhancing your ability to generate clear and actionable insights from visual content.


1. In PowerPoint, open the **Mystic Spice Premium Chai Market Analysis Presentation.pptx** file located in the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder.


    In this simple prompt, you start with the basic **Goal**: _to summarize a PowerPoint._ However, there's no information about why the presentation needs to be summarized or what the summary is needed for.

    Begin with a basic prompt by clearly stating your goal.

    ```
    Summarize this PowerPoint presentation.
    ```

    To improve the prompt, make it a good prompt by adding context. This helps Copilot understand the purpose of the summary.
    For instance: 

    ```
    Summarize this presentation for my boss, including an overview of the main points before meeting with their client.
    ```

    Make it a better prompt by specifying the source or section of the content to focus on.
    For example: 

    ```
    Summarize slides 5 through 10 in this PowerPoint presentation.
    ```

    For the best prompt, set clear expectations about the format and tone.
    For example: 

    ```
    Please format the main points as a bulleted list and use a professional tone.
    ```

    Enter the below crafted prompt:

    ```
    Summarize slides 5-10 in this PowerPoint presentation for my boss that includes an overview of the main points before meeting with their client. Please format the main points as a bulleted list and use a professional tone.
    ```

    In this prompt, the **Goal**, **Context**, **Source**, and **Expectations** are all provided, giving Copilot enough direction to generate a response that meets your needs.


### Exercise 3: Spot trends and visualize data with Copilot in Excel

In this exercise, you'll use Microsoft 365 Copilot in Excel to analyze product sales data and identify key trends. You'll learn how to structure prompts that specify the data scope, set clear analysis goals, and define expected output—enabling Copilot to deliver targeted summaries, insights, and visualizations based on your Excel tables.

To use Copilot in Excel, your data will need to be formatted in one of the following ways:

- As an Excel Table
- As a support Range


1. In Excel, open the **Contoso Chai Tea market trends 2023.xlsx** spreadsheet located in the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder. 

1. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

    In this simple prompt, you start with the basic **Goal**: _to analyze an Excel table._ However, there's no information about why the table needs to be summarized or what the summary is needed for.

    Start with a basic prompt by clearly stating your goal.
    For example: 

    ```
    Analyze this table in Excel.
    ```

    To create a good prompt, add context to help Copilot understand why you need the analysis.
    You might say: 

    ```
    We're looking for the top-selling products from May through August for artisanal chai sales or premade chai sales.
    ```

    Make it a better prompt by specifying the source or scope of the data.
    For example: 

    ```
    Focus on the data from May through August, and only include artisanal and premade chai sales.
    ```

    Finally, to make it the best prompt, set clear expectations for how the results should be delivered.
    Enter the below crafted prompt:

    ```
    Analyze this table in Excel. We're looking for the top selling products from May through August for artisanal chai sales or premade chai sales. Please summarize the top selling product for each month.
    ```

    This prompt gives Copilot everything it needs to come up with a good answer, including the **Goal**, **Context**, **Source**, and **Expectations**.

### Explore more

Try out the final crafted prompt and others with your own Excel table. Here are some suggestions for prompts you might want to try. Copy them and add **Context**, **Sources**, and **Expectations**.  

```
Plot sales by category over time.
```

```
Show total sales for each product.
```

```
Show the total of advertising sales for each region last year.
```