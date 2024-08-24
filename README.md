# Proofread Word Add-In

After setup [Advanced RAG service](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/exploring-the-advanced-rag-retrieval-augmented-generation/ba-p/4197836), users can use this word add-in to proofread content easily:

<img width="480" alt="image" src="https://github.com/freistli/ProofreadAddin/assets/8623897/993d0290-bb65-4eaa-9034-36ded8366897">



https://github.com/freistli/ProofreadAddin/assets/8623897/4300701b-bad7-403f-9b30-d81e00fca35b



# Quick Start

## Publish Proofread Addin as Static Web App

git clone this project. Follow steps in:

https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code

## Modify Manifest.xml

Get the published static web site url, to replace https://localhost:3000 in the manifest.xml.

## Side load manifest.xml to your Word

https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web

- Open new word doc from web portal

  https://www.office.com/launch/Word/

- Create a new Word doc, click File

<img width="224" alt="image" src="https://github.com/freistli/ProofreadAddin/assets/8623897/668b9780-b669-4669-b386-72945c2a9e9f">

- Click Get Add-In -> More Add-ins

<img width="224" alt="image" src="https://github.com/freistli/ProofreadAddin/assets/8623897/e57716da-67ae-4368-95e7-d212f5fd96eb">

- Click My ADD-INS -> Upload My Add-in, upload the manifest.xml

<img width="350" alt="image" src="https://github.com/freistli/ProofreadAddin/assets/8623897/d791299c-706e-49e2-a7d2-2f451f3c7a92">

- In the Word interface, click Show TaskPane

<img width="224" alt="image" src="https://github.com/freistli/ProofreadAddin/assets/8623897/751892dc-8ecc-445b-ae27-3e190f9fabaf">
