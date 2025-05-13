
# Course Title: Craft effective prompts for Microsoft 365 Copilot

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

    [![Screenshot of add new folder](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/add_new.png)](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/add_new.png#lightbox)

4. In **File Explorer**, select **Desktop** and open the **CopilotDocs** folder and click Upload again to upload the folder to OneDrive.

5. Leave **Edge** open and move on to the next task.



# Simplify and extract key information with Copilot in Word

Open your  **OneDrive** folder on your local computer. You should see the files you uploaded in the previous exercise in the folder. 

Open Word in your lab environment. When prompted to sign in, use the **Microsoft Cloud** credentials located in the lab guide's **Environment** tab. When prompted, select **No, this app only**.

Once you are signed into Word, click **Blank document**.

To start using Microsoft 365 Copilot in Word, you can open the **Copilot** pane by selecting the Copilot icon on the far right of the ribbon's **Home** tab.

![Screenshot of the Copilot icon in the Word ribbon.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-ribbon-word.png)

This helpful feature provides answers to questions—broad or specific—about your document. Have a back-and-forth discussion to iterate and refine your results, get a summary or specific information about the document content, or ask it to generate ideas, tables, or lists that you can copy and insert into your document.

![Screenshot of the Copilot panel in Word upon first opening.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-pane-word.png)

In the following example, we turn a basic prompt for Copilot in Word into a well-constructed, contextual prompt that gives you just what you need, in just the way you need it.

## Let's get crafting

In Word, open the **Market Analysis Report for Mystic Spice Premium Chai Tea.docx** document within the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

Here is our first starting prompt:

```
Summarize this Word document.
```

In this simple prompt, you start with the basic **Goal**: _to summarize a Word document._ However, there's no information about why the document needs to be summarized or what the summary is needed for.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **Summarize this Word document.** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand the purpose of the summary and tailor the response accordingly. _"with a brief overview of the main points to discuss with my team during tomorrow's Sales meeting."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot understand which document or part needs to be summarized and provide a more accurate response. _"...the section on Competitive Analysis..."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how to format the summary and what level of detail is required. _"Please keep the summary to 5 key points and use simple language."_ |

Here, we have a better crafted prompt: 

```
Summarize the section on Competitive Analysis in this Word document with a brief overview of the main points to discuss with my team during the tomorrow's Sales meeting. Please keep the summary to 5 key points and use simple language.
```

This prompt has all the details it needs - **Goal**, **Context**, **Source**, and **Expectations** - so Copilot can give you the answer you're looking for.

## Explore more

Try out the final prompt we crafted, but using your own Word document. Customize the **Context**, **Sources**, and **Expectations** so that you get what you need from the document, without any extra stuff you don't.

What are some other ways you can think of to add context or sources or expectations to your prompt? Can you think of other prompting strategies you could use to generate the desired response?


# Identify key information and summarize with Copilot in PowerPoint

Microsoft 365 Copilot in PowerPoint is an AI-powered feature that can help you create, design, and format your slides.  You can type in what you intend to convey with your presentation, and Copilot helps you get it done.

Copilot can help you move past that initial blank slide and get you moving in the right direction. To start using Copilot in PowerPoint, you can open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab.

![Screenshot of the Copilot icon in the PowerPoint ribbon.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-ribbon-powerpoint.png)

From the Copilot pane, you can ask to have the presentation summarized or ask questions about the content on the slides. In the following example, we start with a basic request to summarize the presentation and add other elements to make the prompt more robust.

![Screenshot of the Copilot panel in PowerPoint upon first opening.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-pane-powerpoint.png)

## Let's get crafting

In PowerPoint, open the **Mystic Spice Premium Chai Market Analysis Presentation.pptx** file located in the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder.

Enter the below prompt:

```
Summarize this PowerPoint presentation.
```

In this simple prompt, you start with the basic **Goal**: _to summarize a PowerPoint._ However, there's no information about why the presentation needs to be summarized or what the summary is needed for.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **Summarize this PowerPoint presentation.** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand the purpose of the summary and tailor the response accordingly. _"...for my boss that includes an overview of the main points before meeting with their client."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot understand which presentation or part needs to be summarized and provide a more accurate response. _"...slides 5-10 in this PowerPoint presentation..."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how to format the summary and what level of detail is required. _"Please format the main points as a bulleted list and use a professional tone."_ |

Enter the below crafted prompt:

```
Summarize slides 5-10 in this PowerPoint presentation for my boss that includes an overview of the main points before meeting with their client. Please format the main points as a bulleted list and use a professional tone.
```

In this prompt, the **Goal**, **Context**, **Source**, and **Expectations** are all provided, giving Copilot enough direction to generate a response that meets your needs.

## Explore more

Try out the final prompt we crafted, but using your own PowerPoint presentation. Customize the **Context**, **Sources**, and **Expectations** so that you get what you need from the presentation, without any extra stuff you don't.


# Spot trends and visualize data with Copilot in Excel

Microsoft 365 Copilot in Excel helps you do more with your data in Excel tables by generating formula column suggestions, showing insights in charts and PivotTables, and highlighting interesting portions of data.

In Excel, select **Copilot** on the ribbon to open the chat pane.

![Screenshot of the Copilot icon in the Excel ribbon.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-ribbon-excel.png)

To use Copilot in Excel, your data will need to be formatted in one of the following ways:

- As an Excel Table

- As a support Range

You can create a table, or you can convert a range of cells into a table if you have a data range by following these steps:

1. Select the cell or the range in the data.

1. Select **Home > Format as Table**.

1. In the **Format as Table** dialog box, select the checkbox next to **My table has headers** if you want the first row of the range to be the header row.

1. Select **OK**.

If you prefer to keep your data in a range and not convert it to a table, it will need to meet all of the following requirements:

- Only one header row
- Headers are only on columns, not on rows
- Headers are unique; no duplicate headers
- No blank headers
- Data is formatted in a consistent way
- No subtotals
- No empty rows or columns
- No merged cells

In the following example, we begin with a basic request to analyze a table and progressively add elements to make the prompt more robust.

![Screenshot of the Copilot panel in Excel upon first opening.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/summarize_copilot-pane-excel.png)

## Let's get crafting

In Excel, open the **Contoso Chai Tea market trends 2023.xlsx** spreadsheet located in the **OneDrive\\CopilotDocs\\M01-Summarize-simplify-information-with-microsoft-copilot-microsoft-365** folder. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

Enter the below prompt:

```
Analyze this table in Excel.
```

In this simple prompt, you start with the basic **Goal**: _to analyze an Excel table._ However, there's no information about why the table needs to be summarized or what the summary is needed for.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **Analyze this table in Excel.** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand the purpose of the analysis and adjust the response accordingly. _"We're looking for the top-selling products from May through August for artisanal chai sales or premade chai sales."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot narrow down the scope by telling it to use specific information or ranges. _"...from May through August for artisanal chai sales or premade chai sales...."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how to format the summary and what level of detail is required. _"Please summarize the top-selling product for each month."_ |

Enter the below crafted prompt:

```
Analyze this table in Excel. We're looking for the top selling products from May through August for artisanal chai sales or premade chai sales. Please summarize the top selling product for each month.
```

This prompt gives Copilot everything it needs to come up with a good answer, including the **Goal**, **Context**, **Source**, and **Expectations**.

## Explore more

Try out the final crafted prompt and others with your own Excel table. Here are some suggestions for prompts you might want to try. Copy them and add **Context**, **Sources**, and **Expectations**.  

- Plot sales by category over time.

- Show total sales for each product.

- Show the total of advertising sales for each region last year.

