# Course Title: MS-4005: Craft effective prompts for Microsoft 365 Copilot

## Lesson 2: Creating information with Microsoft Copilot

In this lesson, you learn how to craft effective, contextual prompts for Microsoft 365 Copilot to **create** information in Microsoft 365 apps. If you're looking to enhance your prompting skills for Copilot, this module equips you with the knowledge and techniques to craft prompts that yield accurate and helpful results from Microsoft 365 Copilot.

The focus of this module is **drafting**, **creating**, and **building** information using Microsoft 365 Copilot in different Microsoft 365 apps such as **Word**, **PowerPoint**, **Teams**, **Outlook**, and others. You learn how to use these basic creation capabilities, but also how to craft an efficient prompt that contains all the elements to generate the desired results.


## Lab 


By the end of this lesson, you'll be able to:

1. Identify the key elements of an effective prompt and apply them to your own prompts.

1. Use Copilot to create new agendas, to-do lists, project plans, and more from Word, Excel, and OneNote.

1. Ask Copilot in Outlook to draft new emails, compose replies, and plan meetings.

1. Prompt Microsoft 365 Copilot to generate new ideas, new content, and FAQs from existing files.






# Draft cover letters, marketing plans, and outlines with Microsoft 365 Copilot in Word

To start using Microsoft 365 Copilot  in Word, you can open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab or start writing right within the document. When prompted to login, use the cloud credentials on the environment tab and follow the prompts to complete activation. 

![Screenshot of the Copilot icon in the Word ribbon.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-ribbon-word.png)

To begin drafting right in the body of the document:

1. Open Microsoft Word and start a new blank document.

1. Type or paste your prompt into the **Draft with Copilot** box.

1. Select **Generate**, and Copilot drafts new content for you.

![Draft with Copilot](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/DraftWithCopilot.png)

Once Copilot generates content, select **Keep it** to keep the content, **Regenerate** to regenerate a response, **Discard** to discard the content, or fine tune the draft by entering details into the compose box, like "_Make it more concise_."

![Screenshot of the options bar after using Draft with Copilot in Word.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-prompt-box-word.png)

In the following example, we turn a basic prompt for Copilot in Word into a well-constructed, contextual prompt that gives you just what you need, in just the way you need it.

## Let's get crafting

In Word, open the **Contoso CipherGuard Product Specification.docx** document located in the **OneDrive\\CopilotDocs\\M02-Create-draft-content-with-microsoft-copilot-microsoft-365** folder. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

Enter the below prompt:

```
Draft a marketing proposal.
```

In this simple prompt, you start with the basic **Goal**: _to create a new marketing proposal._ However, there's no information about what the proposal is funding or who is involved.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **_Draft a marketing proposal._** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand what kind of document you want to create and what it will be used for. _"...for Contoso's latest product: CipherGuard. We need to generate three ideas for a marketing campaign..."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot know where to look for specific information. _"...using the product specifications and requirements."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how you want the document to be written and formatted. _"Please include a brief overview of the product, pros and cons for each idea, and ROI projection. Please keep the document to two pages and use optimistic and convincing language."_ |

Enter the below crafted prompt:

```
Draft a marketing proposal for Contoso's latest product: CipherGuard. We need to generate three ideas for a marketing campaign using the product specifications and requirements. Please include a brief overview of the product, pros and cons for each idea, and ROI projection. Please keep the document to two pages and use optimistic and convincing language.
```

Review the results of your prompt and follow up with any questions or refinements, then add it to the end of the document in a new section. **Save** the file so it can be used later.

This prompt gives Copilot everything it needs to come up with a good answer, including the **Goal**, **Context**, **Source**, and **Expectations**.

### Referencing sources

If you want Copilot to base your new document off a file you already have, you can tell it to do that. In the **Draft with Copilot** dialog, select **Reference your content** to choose **_up to 3 files_** that Copilot should look at when creating your new document.

![Reference your content](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/ReferenceYourContent.png)

In the compose box, you can also enter "/" and the name of the file you'd like to reference, which updates the file options shown in the menu for selection.

> **Note:** You must have permission to access the files you're referencing, whether they're located in your organization's SharePoint or OneDrive and can be either Word or PowerPoint files.

## Explore more

Want to give it a try? Use the prompt we crafted with your own documents and presentations. After that, here are some suggestions for other prompts you might want to try.

- _Write an article on the importance of creating work/life balance_.

- _Write a white paper about project management_.

- _Write a job offer letter for a sales position at Contoso. The start date is August 1, and the salary is $60,000 per year plus bonuses_.

# Build new slides, agendas, and to-do lists with Microsoft 365 Copilot in PowerPoint

Microsoft 365 Copilot  in PowerPoint is an AI-powered feature that can help you create, design, and format your slides.  You can type in what you intend to convey with your presentation, and Copilot helps you get it done.

Copilot can help you move past that initial blank slide and get you moving in the right direction. To start using Copilot in PowerPoint, you can open the **Copilot** pane via the Copilot icon in the ribbon's **Home** tab.

![Screenshot of the Copilot icon in the PowerPoint ribbon.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-ribbon-powerpoint.png)

From the Copilot pane, you can begin creating a new presentation from a Word document or about a desired topic. In the example, we start with a basic request to create a presentation about a topic and add other elements to make the prompt more robust.

![Screenshot of the Copilot panel in PowerPoint upon first opening.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-pane-powerpoint.png)

> **Note:** Currently, Copilot in PowerPoint is only able to create presentations from Word documents.

## Let's get crafting

<!-- If you haven't done so yet, download the following files and save the file to your **OneDrive folder** so they appear in your MRU list:

- **_[Market Trend Report- Protein shake.docx](https://go.microsoft.com/fwlink/?linkid=2268827)_** -->

Open PowerPoint then open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along.

<!-- In Word, open the **Market Trend Report- Protein shake.docx** document located in the **OneDrive\\CopilotDocs\\M02-Create-draft-content-with-microsoft-copilot-microsoft-365** folder. Open the **Copilot** pane by selecting the Copilot icon in the ribbon's **Home** tab. Enter the prompts below and follow along. -->

Enter the below prompt:

```
Create a new PowerPoint presentation.
```

In this simple prompt, you start with the basic **Goal**: _to build a new PowerPoint presentation._ However, there's no information about what the presentation is about or what it should look like.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **_Create a new PowerPoint presentation._** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand what kind of document you want to create and what it will be used for. _"We need to present the product's features and benefits to potential clients."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot know where to look for specific information. _"...using the latest from **/Market Trend Report- Protein shake.docx**."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how you want the document to be written and formatted. _"Please include an overview of the product, its key features and benefits, and a comparison to similar products in the market. Please use simple language."_ |

Enter the below crafted prompt:

```
Create a new PowerPoint presentation using the latest from Market Trend Report- Protein shake.docx. We need to present the product's features and benefits to potential clients. Please include an overview of the product, its key features and benefits, and a comparison to similar products in the market. Please  use simple language.
```

With the **Goal**, **Context**, **Source**, and **Expectations** all laid out, Copilot has everything it needs to give you a great response.

### Referencing sources

As in the example, if you want Copilot to base your new presentation off a file you already have, you can tell it to do that. In the prompt window, select **Create presentation from file** to choose **_up to 3 files_** that Copilot should look at when creating your new document.

In the compose box, you can also enter "/" and the name of the file you'd like to reference, which will update the file options shown in the menu for selection.

>**Note:** You must have permission to access the files you're referencing, whether they're located in your organization's SharePoint or OneDrive and can be either Word or PowerPoint files.

### Best practices when creating a presentation from a Word document

Leverage **Word Styles** to help Copilot understand the structure of your document. By using **Styles** in Word to organize your document, Copilot will better understand your document structure and how to break it up into slides of a presentation. Structure your content under **Titles** and **Headers** when appropriate and Copilot will do its best to generate a presentation for you.

### Include images that are relevant to your presentation

When creating a presentation, Copilot will try to incorporate the images in your Word document. If you have images that you would like to be brought over to your presentation, be sure to include them in your Word document.

### Start with your organization’s template

If your organization uses a standard template, start with this file before creating a presentation with Copilot. Starting with a template will let Copilot know that you would like to retain the presentation’s theme and design. Copilot will use existing layouts to build a presentation for you.

> **Note:** This feature is available to customers with a Microsoft 365 Copilot license or Copilot Pro license. For more information, see [Create a presentation from a file with Copilot](https://support.microsoft.com/office/create-a-new-presentation-3222ee03-f5a4-4d27-8642-9c387ab4854d).


# Draft emails, replies, and meeting agendas with Microsoft 365 Copilot in Outlook

Copilot in Outlook makes inbox management easier with AI-powered assistance to help you write emails quickly and turn long email threads into short summaries. It combines the power of large language models (LLMs) with Outlook data to help you stay productive in the workplace. It can summarize email threads (also known as conversations), pulling out key points from multiple messages.

<!-- Copilot in Outlook can help you quickly draft an email or a reply to an existing conversation.

1. In Outlook, select **Home > New Mail > Mail**.

1. To start a new message, select the **Copilot icon** from the toolbar.

1. Select **Draft with Copilot** from the drop-down menu.

    ![Screenshot of the Copilot icon in the Outlook toolbar.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-toolbar-outlook.png)

1. In the Copilot box, **type your prompt**.

1. Select **Generate options** to choose your desired length and tone.

    ![Screenshot of the available options to customize your draft in Copilot in Outlook.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-generate-options-outlook.png)

1. When finished, select **Generate**. Copilot drafts a message for you.

1. Review the message. If it’s not quite what you want, choose **Regenerate draft** and Copilot create a new version.

1. To start over, change your prompt and select **Generate** again.

1. Once satisfied with the result, select **Keep**.

1. Edit the draft as needed, and then select **Send**.

    ![Screenshot of a draft email generated by Copilot in Outlook.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-draft-results-outlook.png) -->

## Let's get crafting

Open Outlook.

Click **New mail** then expand Copilot on the right and click **Draft**. 

In the box that appears, enter the below prompt:

```
Draft a new email.
```

![Draft new email](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/DraftNewEmailSimple.png)

In this simple prompt, you start with the basic **Goal**: _to draft a new email._ However, there's no information about what the email will even be about, who to send it to, or what you want it to sound like.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **_Draft a new email._** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand what the email should be about and who the audience is. _"...to my client, Allan Deyoung, regarding the status of their support ticket."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot know where to look for specific information. _"Refer to the latest update from my notes: The issue has been escalated to Tier 2 support, and a resolution is expected within 48 hours."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how you want the document to be written and formatted. _"The email should sound professional and technical, but written with empathy."_ |

Click **Discard**, then expand Copilot on the right and click **Draft** again.

In the box that appears, enter the below crafted prompt:

```
Draft a new email to my client, Allan Deyoung, regarding the status of their support ticket. Refer to the latest update from my notes: The issue has been escalated to Tier 2 support, and a resolution is expected within 48 hours. The email should sound professional and technical, but written with empathy.
```

In this prompt, Copilot has all the information it needs to give you a solid answer, thanks to the **Goal**, **Context**, **Source**, and **Expectations** in this prompt.

# Brainstorm new ideas, lists, and reports with Microsoft 365 Copilot Chat

Microsoft 365 Copilot Chat (Copilot Chat) combines the power of artificial intelligence (AI) with your work data and apps to help you unleash creativity, unlock productivity, and uplevel skills. It works across multiple apps and content, giving you the power of AI together with your secure work data. Its ability to synthesize information and create things from multiple sources at once empowers you to tackle broader goals and objectives.

To compare, Copilot in the different Microsoft 365 apps (such as Word or PowerPoint) is specifically orchestrated to help you **within that one app**. For example, Copilot in Word is designed to help you better draft, edit, and consume content. In PowerPoint, it’s there to help you create better presentations. But with Copilot Chat, we can pull that all together into a new experience.

You can access Copilot Chat in several ways:

- Use Copilot in desktop and mobile versions of Microsoft Teams. See [Use Copilot Chat in Teams](https://support.microsoft.com/topic/open-microsoft-365-chat-in-teams-c6de0a62-4f9e-479d-b5f2-af036e342181).

- Access Copilot Chat at Microsoft.com/Copilot. See [Use Copilot Chat at Microsoft365.com/copilot](https://support.microsoft.com/topic/use-microsoft-365-chat-at-microsoft365-com-or-in-the-microsoft-365-office-app-4a2538f9-962f-4c7c-a368-f6006bc13d6f).

![Screenshot of the Copilot Chat experience in Microsoft Teams.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-chat-experience-teams.png)

## Let's get crafting

Navigate to the below url to access CoPilot Chat. Enter the **Microsoft Cloud** credentials located in the lab guide's **Enironment** tab if prompted. Follow along with the prompts.

```
Microsoft365.com/copilot
```

Follow along with the prompts using **Contoso CipherGuard Product Specification.docx**.

Enter the below prompt:

```
Build a meeting agenda.
```

In this simple prompt, you start with the basic **Goal**: _to build out a meeting agenda_. However, there's no information about the purpose or goal of the meeting.

| Element | Example |
| :------ | :------- |
| **Basic prompt:** Start with a **Goal** | **_Build a meeting agenda._** |
| **Good prompt:** Add **Context** | Adding **Context** can help Copilot understand why you are calling the meeting and what you want to discuss. _"...for a client meeting to last an hour, including the project goals, mission statement, and scheduled completion date."_ |
| **Better prompt:** Specify **Source(s)** | Adding **Sources** can help Copilot know where to look for specific information. _"Use information from **/Contoso CipherGuard Product Specification.docx** and look for open items and unanswered questions."_ |
| **Best prompt:** Set clear **Expectations** | Lastly, adding **Expectations** can help Copilot understand how you want the document to be written and formatted. _"The agenda should be in a table format with time allotments, and make sure to give people the chance to ask questions at the end."_ |

Enter the below crafted prompt:

```
Build a meeting agenda for a client meeting to last an hour, including the project goals, mission statement, and scheduled completion date. Use information from Contoso CipherGuard Product Specification.docx and look for open items and unanswered questions. The agenda should be in a table format with time allotments, and make sure to give people the chance to ask questions at the end.
```

![Screenshot the crafted prompt results against the sample document using Copilot Chat.](https://labstoragesupport.blob.core.windows.net/lessoncontent/ms-4005-ai-led/media/create_copilot-chat-draft-agenda-teams.png)

Review the agenda and make any adjustments or refinements, then add it to your meeting invite in Teams.

### Referencing sources

As in the example, if you want Copilot to base your new presentation off a file, meeting, or person (even all three), you can tell it to do that. In the prompt window, just start typing a forward slash "/" and a popout window will offer you recent meetings, files, or people to reference.

>**Note:** You must have permission to access the files you're referencing, whether they're located in your organization's SharePoint or OneDrive and can be either Word, Excel, or PowerPoint files.

## Explore more

Here are some suggestions for prompts you might want to try. Copy them or modify them to suit your needs.

- What happened in my last meeting?

- Catch up on unread chats.

- Draft a message that OKRs are due next week.

- Tell my team how we updated the product strategy.

- Summarize the chats, emails, and documents about [a customer] escalation that happened last night.

- What is the next milestone on [a project]? Are there any risks? Help me brainstorm a list of some potential ways to mitigate.

- Write a planning overview in the style of [a file] that contains the timeline from [a different file] and incorporates the project list in the email from [a person].
