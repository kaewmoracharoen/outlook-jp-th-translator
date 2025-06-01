// ตรวจสอบให้แน่ใจว่า Office.js โหลดเสร็จแล้ว
Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // โค้ดที่เกี่ยวข้องกับ Outlook จะอยู่ที่นี่
        document.getElementById("btnGetEmailContent").onclick = getEmailContent;
        document.getElementById("btnTranslateJPtoTH").onclick = function() { translateText('jp_to_th'); };
        document.getElementById("btnTranslateTHtoJPEN").onclick = function() { translateText('th_to_jp_en'); };
        
        console.log("Outlook Add-in is ready!");
    }
});

function getEmailContent() {
    // Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(asyncResult) {
    //     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //         document.getElementById("originalText").value = asyncResult.value;
    //         document.getElementById("inputText").value = asyncResult.value; // ใส่ในช่อง input ด้วย
    //         console.log("Email content retrieved.");
    //     } else {
    //         console.error("Failed to get email body: " + asyncResult.error.message);
    //         document.getElementById("originalText").value = "Error: " + asyncResult.error.message;
    //     }
    // });
    // **หมายเหตุ:** การดึงเนื้อหาอีเมลจริง จะต้องมีการให้สิทธิ์และทดสอบอย่างละเอียด
    // ในขั้นตอนนี้ เราจะใส่ข้อความตัวอย่างไปก่อน
    document.getElementById("originalText").value = "นี่คือเนื้อหาอีเมลตัวอย่าง (ตัวจริงจะถูกดึงเมื่อเชื่อมต่อ Backend)";
    document.getElementById("inputText").value = "นี่คือเนื้อหาอีเมลตัวอย่าง (ตัวจริงจะถูกดึงเมื่อเชื่อมต่อ Backend)";
    console.log("Placeholder email content set.");
}

function translateText(direction) {
    var inputText = document.getElementById("inputText").value;
    if (!inputText.trim()) {
        document.getElementById("translatedText").value = "กรุณาป้อนข้อความที่ต้องการแปล";
        return;
    }

    // **ส่วนนี้คือส่วนที่จะต้องเรียก Backend Service ของคุณในอนาคต**
    // ในตอนนี้ เราจะแค่แสดงข้อความจำลองการแปล
    let translatedOutput = "";
    if (direction === 'jp_to_th') {
        translatedOutput = `[ผลการแปล ญี่ปุ่น -> ไทย ของ: "${inputText}"]\n(เชื่อมต่อ Backend เพื่อแปลจริง)`;
    } else if (direction === 'th_to_jp_en') {
        translatedOutput = `[ผลการแปล ไทย -> ญี่ปุ่น/อังกฤษ ของ: "${inputText}"]\n(เชื่อมต่อ Backend เพื่อแปลจริง)`;
    }
    
    document.getElementById("translatedText").value = translatedOutput;
    console.log(`Simulated translation for direction: ${direction}`);
}