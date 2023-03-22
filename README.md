# Refiny: Polish text with OpenAI's GPT 3.5 &nbsp;
Adding OpenAI's text polishing capability to Microsoft Word
<br /><br />

![Alt Text](assets/images/demo.gif)

![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)&nbsp;&nbsp;
![Microsoft Word](https://img.shields.io/badge/Microsoft_Word-2B579A?logo=microsoft-word&logoColor=white)&nbsp;&nbsp;


## 1. About
Improve readability of the selected text in a clear and concise manner with the click of a button in MS-Word. This is a Word macro written in VBA that communicates with the OpenAI Chat Completion API endpoint and uses the gpt-3.5-turbo which is on-par with the text-davinci-003 model for instruction-following tasks.
### OpenAI costs to consider
- Signing up for an OpenAI account provides you with **18$** free credits.
- Average manuscript size: 15,000 words -> 20,480 tokens
- Associated approximate costs: 0.002$/1k tokens x 20,480 -> **approx. 5 cents**

<br />

## 2. Installation
### 2.1 Obtain an OpenAI API Key

Signup for an OpenAI account at <a href="https://platform.openai.com/account/api-keys">platform.openai.com/account/api-keys</a>. Create an API key and copy the key for next steps.
![Alt Text](assets/images/api_key.png)

<br />

### 2.2 Download Folder for Images
All images created will be saved to the following default folder in Windows 10/11
```
C:\Users\Public\Pictures
```


<br />

### 2.3 AI Models

This application uses the following OpenAI model:
 
| Model  | Description | Iteration |
| ------------- | ------------- |------------- | 
| DALL·E  | A model that can generate and edit images given a natural <br /> language prompt. The current model has more realistic, accurate, <br /> and 4x greater resolution images than the original model. | 2nd | 

Further information about all OpenAI models can be found here: <a href="https://platform.openai.com/docs/models/overview">https://platform.openai.com/docs/models/overview</a>

We endeavour to test and integrate newer models when they are become Generally Available (GA). Models released as a 'Limited Beta' will not be integrated until they become GA.

<br />

### 2.4 OpenAI Status

Should you receive error messages for a period of 5 minutes or more, please check the Labs Uptime status at the following link: <a href="https://status.openai.com/">https://status.openai.com/</a>

<br /><br />
    
<!-- INSTRUCTIONS -->
## 3. Instructions

  - Write a desciption of the image you want DALL·E to create anywhere in Microsoft Word
  - Click the **Image Generation** button in the AI Assistant tab
  - Wait a few seconds for DALL·E to respond
  - The bespoke image will appear under the initial text descrption
<br /><br />

### 3.1 Image Generation Example 1

User Input
```
Watercolor painting house on the water with trees
```

DALL·E Output

<img src=".github/assets/images/user-example-1.png" width=100% height=100%>
<br />

### 3.2 Image Generation Example 2

User Input
```
An astronaut riding a horse in a photorealistic style
```

DALL·E Output

<img src=".github/assets/images/user-example-2.png" width=100% height=100%>
<br />

### 3.3 Top 10 Image Creation Ideas
If you are stuck for inspiration try one of these creation prompts:
```
1. Oil painting mountains and trees
2. Watercolor house by the water
3. Pointillism landscape
4. Futuristic cars
5. Mona Lisa laughing
6. An astronaut riding a horse in a photorealistic style
7. Fantasy artwork
8. Alien riding an elephant in space holding a phone
9. Painting of a kangaroo cartoon style
10. Bowl of Fruit, Violin and Bottle cubism style
```
<br />




<!-- Installation -->
## 4. Installation

There are 4 basic steps in order to add a ChatGPT button into Microsoft Word:
  1. Enable the Developer Tab
  2. Import the VBA script file
  3. Create the **Image Generation** button 
  4. Add your OpenAI APIKey
  

Each of these steps are fully outlined below. 
<br /><br />

### 4.1 Enable the Developer Tab

The Developer tab isn't displayed by default, but you can add it to the ribbon.

**Step 1** - On the File tab, go to Options > Customize Ribbon.

**Step 2** - Under Customize the Ribbon and under Main Tabs, select the Developer check box.

<img src=".github/assets/images/enable-developer-tab-highlighted.png" width=100% height=100%>
<br />

The latest instructions to enable the Developer Tab from Microsoft can be found here: 
<a href="https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792">https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792</a>
<br /><br />

### 4.2 Import the ChatgptImageGeneration.bas file

**Step 1** - Download and Save the latest ChatgptImageGeneration.bas file from this repository.
<br />

Keep the location of where the file is saved as you will need it later.<br />
<br />

**Step 2** - On the Developer tab, click the Visual Basic button.

<img src=".github/assets/images/developer-tab-visual-basic.png" width=100% height=100%>
<br />

**Step 3** - On the File tab, go to Import File...

<img src=".github/assets/images/visual-basic-file-import-section.png" width=100% height=100%>
<br />

**Step 4** - Select the ChatgptImageGeneration.bas file and click Open
<br /><br />

### 4.3 Add your ChatGPT button into the Microsoft Word Ribbon

**Please Note:** This project closely relates to our *Add ChatGPT to Microsoft Word* project found here: <a href="https://github.com/analyticsinmotion/add-chatgpt-to-microsoft-word">https://github.com/analyticsinmotion/add-chatgpt-to-microsoft-word</a>. If you have have already added that project into Microsoft Word you can start at Step 4 of this section.
<br /><br />

**Step 1** - Add a new tab **(OPTIONAL)**
<br />
  - On the File tab, go to Options > Customize Ribbon
  - Click New Tab
<br />

<img src=".github/assets/images/options-customize-ribbon-new-tab.png" width=40% height=40%>
<br />

 **Step 2** - Rename the New Tab to **AI Assistant** **(OPTIONAL)**

<img src=".github/assets/images/options-customize-ribbon-rename-tab.png" width=35% height=35%>
<br />

**Step 3** - Rename New Group (Custom) to **ChatGPT** **(OPTIONAL)**

<img src=".github/assets/images/rename-new-group.png" width=35% height=35%>
<br />

**Step 4** - Select Macros in the Choose Commands from dropdown box

<img src=".github/assets/images/choose-commands-from-macros.png" width=35% height=35%>
<br />

**Step 5** - Select the ChatgptImageGeneration Macro and click Add >>

<img src=".github/assets/images/add-the-macro-into-new-group.png" width=75% height=75%>
<br />

**Step 6** - Rename button to **Image Generation**, select a Symbol and click OK

<img src=".github/assets/images/rename-button-to-text-completion.png" width=35% height=35%>
<br />

After the preceding steps have been completed the Microsoft Word screen should look like the following:

<img src=".github/assets/images/screen-after-chatgpt-button-added.png" width=100% height=100%>
<br />



### 4.4 Add your OpenAI APIKey into Windows

**Step 1** - Open the Start menu and start typing "environment variables". When the best match appears click "Edit the system environment variables" result.

<img src=".github/assets/images/add-env-var-step-1.png" width=75% height=75%>
<br />

**Step 2** - Click the "Environment variables" button under the "Advanced" tab.

<img src=".github/assets/images/add-env-var-step-2.png" width=50% height=50%>
<br />

**Step 3** - Create a new user variable by clicking "New" under the "User Variables" section.

<img src=".github/assets/images/add-env-var-step-3.png" width=50% height=50%>
<br />

**Step 4** - Type the variable name **OPENAI_API_KEY** in the first field and your OpenAI APIKEY in the variable value field. Then click OK.

<img src=".github/assets/images/add-env-var-step-4.png" width=50% height=50%>
<br />

**Step 5** - **IMPORTANT** You must restart Windows to apply the new environment variable
<br /><br />


<!-- Best Practices for API Key Safety -->
## 5. Best Practices for API Key Safety

Your OpenAI APIKEY key/s should be kept secure and private at all times.

Please follow the best practices guide for API security from OpenAI 
<br />
<a href="https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety">https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety</a>
