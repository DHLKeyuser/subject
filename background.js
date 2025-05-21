chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {

  if (message.action === "saveShipmentData") {

    const blob = new Blob([message.content], { type: "text/plain" });

    const reader = new FileReader();

    reader.onloadend = () => {

      const dataUrl = reader.result;

      chrome.downloads.download({

        url: dataUrl,

        filename: "shipment_data_temp.txt",

        saveAs: false,

        conflictAction: "overwrite"

      });

    };

    reader.readAsDataURL(blob);

  }

});
 