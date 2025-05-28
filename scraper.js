(() => {

  if (!location.href.includes("SendungBearbeitung?")) return;

  try {

    const shipmentID = Array.from(document.querySelectorAll('#TableLegs a'))

      .map(a => a.textContent.trim())

      .find(text => /^4\d{7}$/.test(text)) || "";

    const dnField = document.querySelector('#elmtKopf\\.ReferenzReference02')?.value.trim() || "";

    const allDNs = dnField ? [dnField] : [];

    const tableText = document.querySelector('#TableLegs')?.innerText.toLowerCase() || "";

    let forwarder = "";

    if (tableText.includes("fedex") && tableText.includes("freight")) forwarder = "FedEx Freight";

    else if (tableText.includes("fedex"))                     forwarder = "FedEx";

    else if (tableText.includes("dsv"))                       forwarder = "DSV";

    else if (tableText.includes("maersk"))                    forwarder = "MAERSK";

    else if (tableText.includes("tromp"))                     forwarder = "TROMP";

    else if (tableText.includes("dhl") && tableText.includes("global")) forwarder = "DHL-GF";

    else if (tableText.includes("kuehne"))                    forwarder = "KNN";

    // New: only scrape Geschäftsfall if it's NOT the default TRM1

    let businessCaseText = "";

    const bcSelect = document.getElementById("elmtKopf.Geschaeftsfall");

    if (bcSelect) {

      const opt = bcSelect.options[bcSelect.selectedIndex];

      if (opt && opt.value && opt.value !== "TRM1") {

        const prefix = opt.value.replace(/\d+$/, "");

        const desc   = opt.title?.trim() || opt.text.trim();

        businessCaseText = `${prefix}:\n${desc}`;

      }

    }

    if (shipmentID && allDNs.length > 0 && forwarder) {

      let result = 

        `Shipment ID:\n${shipmentID}` +

        `\nDN:\n${allDNs.join("\n")}` +

        `\nForwarder:\n${forwarder}`;

      if (businessCaseText) {

        result += `\n${businessCaseText}`;

      }

      window.name = result;

      chrome.runtime.sendMessage({

        action: "saveShipmentData",

        content: result

      });

    }

  } catch (e) {

    console.error("Scraper error:", e);

    window.name = "Scraping failed: " + e.message;

  }

})();
 