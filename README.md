function sendMessagesFromSheet7() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // ⏳ ينتظر حتى 30 ثانية إذا في كود آخر شغال
  } catch (e) {
    Logger.log("🔁 كود آخر قيد التشغيل، تم تجاهل هذا التشغيل.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الورقة7");
  if (!sheet) return Logger.log("❌ لم يتم العثور على الورقة 7");

  const data = sheet.getDataRange().getValues();
  const now = new Date();

  const messages = [
    "شكراً على زيارتك لـ *لوكيشن كافيه*! 🌟 حابين نسمع رأيك. قيّم تجربتك واحصل على خصم 15% لزيارتك الجاية. اضغط 1 وشاركنا رأيك!",
    "يسعدنا إنك اخترت *لوكيشن كافيه*! 😊 شاركنا تقييمك واستمتع بخصم 15% في زيارتك المقبلة. اضغط 1 وخبرنا عن تجربتك!",
    "أهلًا وسهلًا في *لوكيشن كافيه*! 🌟 قول لنا رأيك واحصل على خصم 15% لزيارتك القادمة. اضغط 1 وخلينا نعرف تجربتك!"
  ];

  const apiUrl = "https://app.arrivewhats.com/api/send";
  const accessToken = "66f85a4411dc4";
  const instanceId = "67AF2CB7C7B5F";

  for (let i = 1; i < data.length; i++) {
    const phone = data[i][0];
    const rawTimestamp = data[i][1];
    const status = data[i][2];

    if (!phone || !rawTimestamp || status === "Sent") continue;

    const timestamp = (rawTimestamp instanceof Date) ? rawTimestamp : new Date(rawTimestamp);
    const timeDiff = now - timestamp;

    if (timeDiff < 10 * 60 * 1000) {
      if (status !== "Processing") {
        sheet.getRange(i + 1, 3).setValue("Processing");
        SpreadsheetApp.flush();
        Logger.log(`⏳ Row ${i + 1}: لم تمر 10 دقائق بعد - الحالة: Processing`);
      }
      continue;
    }

    const message = messages[Math.floor(Math.random() * messages.length)];
    const url = `${apiUrl}?number=${encodeURIComponent(phone)}&type=text&message=${encodeURIComponent(message)}&instance_id=${instanceId}&access_token=${accessToken}`;

    try {
      const response = UrlFetchApp.fetch(url, { method: "get" });
      const json = JSON.parse(response.getContentText());

      if (json.status === "success") {
        sheet.getRange(i + 1, 3).setValue("Sent");
        SpreadsheetApp.flush();
        Logger.log(`✅ Row ${i + 1}: تم الإرسال إلى ${phone}`);
      } else {
        sheet.getRange(i + 1, 3).setValue("Failed - " + json.message);
        SpreadsheetApp.flush();
        Logger.log(`❌ Row ${i + 1}: فشل الإرسال إلى ${phone} - ${json.message}`);
      }

    } catch (err) {
      sheet.getRange(i + 1, 3).setValue("Error - " + err.message);
      SpreadsheetApp.flush();
      Logger.log(`⚠️ Row ${i + 1}: خطأ أثناء الإرسال - ${err.message}`);
    }

    // تأخير عشوائي من 10 إلى 30 ثانية
Utilities.sleep(Math.floor(Math.random() * (15000 - 5000 + 1)) + 5000);
  }

  // 🔓 فك القفل في النهاية
  lock.releaseLock();
}
