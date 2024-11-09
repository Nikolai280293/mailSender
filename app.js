const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");

// Настройка транспорта для отправки почты через SMTP
const transporter = nodemailer.createTransport({
  host: "smtp.yandex.ru",
  port: 465,
  secureConnection: true,
  auth: {
    user: "name",
    pass: "password",
  },
});

const mailOptions = {
  from: "info@1med.tv",
  subject: "Мероприятие НМО. Международная практическая школа «Тренды и традиции в оперативной проктологии. Версия 7.0»",
  html: `
  
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>

<body style="font-family: Arial, sans-serif; margin: 0; padding: 0;">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width: 100%; margin: 0 auto;">
        <tr>
            <td style="padding: 20px; background-color: #ffffff;">
                <h5 style="font-size: 14px; margin: 0 0 10px 0;">Добрый день!</h5>
                <p style="font-size: 14px; margin: 0 0 10px 0;">Вручаем Вам Свидетельство о прохождении обучения в рамках международной практической школы «Тренды и традиции в оперативной проктологии. Версия 7.0», 13 - 14 сентября 2024 г.</p>
                <p style="font-size: 14px; margin: 0 0 10px 0;">Во вложении письма Свидетельство о прохождении обучения с одним индивидуальным кодом подтверждения.</p>
                <p style="font-size: 14px; margin: 0 0 10px 0;">Для активации кредитов необходимо:</p>
                <table cellpadding="0" cellspacing="0" style="font-size: 14px; margin: 0 0 10px 0;">
                    <tr>
                        <td>1.</td>
                        <td>Зарегистрироваться на сайте <a href="http://www.edu.rosminzdrav.ru/" target="_blank">http://www.edu.rosminzdrav.ru/</a></td>
                    </tr>
                    <tr>
                        <td>2.</td>
                        <td>На сайте <a href="http://www.edu.rosminzdrav.ru/" target="_blank">http://www.edu.rosminzdrav.ru/</a>, в Личном кабинете ввести индивидуальный код, указанный в Свидетельстве.</td>
                    </tr>
                </table>
                <p style="font-size: 14px; margin: 0 0 10px 0;">Баллы могут начисляться в течение суток.</p>
                <p style="font-size: 14px; margin: 0 0 10px 0;">При возникновении вопросов Вы можете обращаться в техподдержку Портала непрерывного медицинского и фармацевтического образования Минздрава России <a href="http://www.edu.rosminzdrav.ru/" target="_blank">http://www.edu.rosminzdrav.ru/</a></p>
            </td>
        </tr>
    </table>
</body>

</html>

  `,
  text: "Свидетельство о прохождении обучения",
};

const writeLog = (message) => {
  const timestamp = new Date().toISOString();
  fs.appendFileSync("log.txt", `[${timestamp}] ${message}\n`);
};

const emailList = [];
fs.createReadStream('emails.csv')
  .pipe(csv())
  .on('data', (row) => {
    if (row.email) {
      emailList.push(row.email);
    }
  })
  .on('end', () => {
    console.log("Список email-адресов загружен.");

    const sentEmails = [];
    const failedEmails = [];

    const sendEmail = (email, attachments, retryCount = 0) => {
      transporter.sendMail({
        ...mailOptions,
        to: email,
        attachments,
      }, (error, info) => {
        if (error) {
          console.error(`Ошибка при отправке письма на ${email}:`, error);

          writeLog(`Ошибка при отправке на ${email}: ${error.message}`);

          if (error.responseCode === 450 && retryCount < 12) {
            const delay = (retryCount + 1) * 9000;

            console.log(`Повторная отправка письма на ${email} через ${delay / 1000} секунд (попытка ${retryCount + 1})...`);

            setTimeout(() => {
              sendEmail(email, attachments, retryCount + 1);
            }, delay);
          } else {
            failedEmails.push(email);
          }
        } else {
          if (info && info.messageId) {
            console.log(`Письмо успешно отправлено на ${email}!`);
            writeLog(`Письмо успешно отправлено на ${email}: ${info.messageId}`);
            sentEmails.push(email);
          }
        }
      });
    };

    emailList.forEach((email, index) => {
      const filename = `Сертификат_${index + 1}.pdf`;
      const pathToFile = path.join(__dirname, "resources", filename);

      const attachments = [];
      if (fs.existsSync(pathToFile)) {
        attachments.push({
          filename,
          path: pathToFile,
        });
      } else {
        console.error(`Файл ${filename} не найден!`);
        writeLog(`Файл ${filename} не найден для ${email}`);
        return;
      }
      const delay = index * 15000;
      setTimeout(() => {
        sendEmail(email, attachments);
      }, delay);
    });
    setTimeout(() => {
      const successMessage = `Отправлены письма на следующие адреса: ${sentEmails.join(", ")}`;
      const failureMessage = `Не удалось отправить письма на следующие адреса: ${failedEmails.join(", ")}`;

      console.log(successMessage);
      console.log(failureMessage);

      writeLog(successMessage);
      writeLog(failureMessage);
    }, emailList.length * 15000 + 2000);
  })
  .on('error', (err) => {
    console.error('Ошибка при чтении CSV файла:', err);
    writeLog(`Ошибка при чтении CSV файла: ${err.message}`);
  });