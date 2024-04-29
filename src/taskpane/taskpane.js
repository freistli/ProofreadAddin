/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { client } from "@gradio/client";

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

 
export async function updateIframeSrc() {
    const iframeSrcInput = document.getElementById('iframe-src');
    const iframe = document.getElementById('proofreading-iframe');
    iframe.src = iframeSrcInput.value+"/proofreadaddin/";
    setInLocalStorage("serviceUrl",iframeSrcInput.value); 
    console.log("Save Service URL:" + iframeSrcInput.value);
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("Proofreading").onclick = proofreading;
    document.getElementById("update-iframe-src").onclick = updateIframeSrc;

    const serviceUrl = getFromLocalStorage("serviceUrl");

    console.log("Get Stored Service URL:" + serviceUrl);

    if (serviceUrl) {
      document.getElementById("iframe-src").value = serviceUrl;
      updateIframeSrc();
    }

    window.addEventListener('message', event => {
      console.log("gradio posted message arrived");
      console.log(event.origin);

      if (event.origin === 'https://localhost:3000' || true) {
          
         run().then(() => {console.log("run() finished")});

      } else {
          
          return;
      }   

  });
    
  }
}); 

export async function run() {
  return Word.run(async (context) => {
  
    // insert a paragraph at the end of the document.
    //const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    const selection = context.document.getSelection();

    context.load(selection, 'text');

    context.sync().then(function () {

    const selectedText = selection.text; 

    const frame = document.getElementById('proofreading-iframe');
    frame.contentWindow.postMessage(selectedText, '*');
   
    });

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";

    await context.sync();
  });
}

export async function proofreading() {
  return Word.run(async (context) => {


    const systemMessage = "Criticize the proofread content, especially for wrong words. Only use 当社の用字・用語の基準,  送り仮名の付け方, 現代仮名遣い,  接続詞の使い方 ，外来語の書き方，公正競争規約により使用を禁止されている語  製品の取扱説明書等において使用することはできない, 常用漢字表に記載されていない読み方, and 誤字 proofread rules, don't use other rules those are not in the retrieved documents.                Pay attention to some known issues:もっとも, または->又は, 「ただし」という接続詞は原則として仮名で表記するため,「又は」という接続詞は原則として漢字で表記するため。また、「又は」は、最後の語句に“など”、「等(とう)」又は「その他」を付けてはならない, 優位性を意味する語.               Firstly show 原文, use bold text to point out every incorrect issue, and then give 校正理由, respond in Japanese. Finally give 修正後の文章, use bold text for modified text. If everything is correct, tell no issues, and don't provide 校正理由 or 修正後の文章."

    const app = await client("http://127.0.0.1:7860/");
    
    result = await app.predict("/predict",[
        "rules",	 
        systemMessage,
        document.getElementById("item-content").value]);
    document.getElementById("item-proofread").value   = result
    await context.sync();
  });
}