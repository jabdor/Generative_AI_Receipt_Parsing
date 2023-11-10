# Generative_AI_Receipt_Parsing

Working script to take input receipt images, parse them out using OpenAI's API (specifically the vision api), rename the files with contextual information, and then upload the line items to a google spreadsheet.

The parsing extracts header and line-item information:

  Header:
    Vendor
    Receipt Date
    Total Receipt Cost
  Line-Item:
    Line Item
    Line Item Cost
    Line Item Description
    Line Item Category

On average, the token cost is around 1000 tokens (mixed with the request and response). At this point in time, it's about .01 - .03 per image. 

## Features

- Converts PDFs to jpgs. The API requires images to be in jpg or png format.
- Batches through images in a source directory
- Outputs the ChatGPT analysis as a python directory for possible future use cases
- Renames files in a destination folder with the header information
- Exports line items to a google spreaddsheet
- Some light error handling (more needed). Handling in place for when receipts are blurry or don't have line items.

## Prerequisits

- OpenAI still has frequent outages. This causes the script to stall when there is an outage. They're not common, but it can occur. Here's the status page to check: https://status.openai.com/
- You will need an OpenAI API key, and you will need to have loaded credit to your account.
- You will need to go through the proper setup for the google sheet api. The main steps are to create the proper google project, create a service account, and then generate a 'credentials.json' file.

## Future Ideas

- Add better logging
- Create an executable script
- Perform a status page check prior to starting the script
