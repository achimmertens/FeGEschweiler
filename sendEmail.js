const nodemailer = require('nodemailer');
const config = require('./config');

let transporter = nodemailer.createTransport({
    host: 'mail.gmx.net',
    port: 587,
    secure: false,
    auth: {
        user: config.username,
        pass: config.password
    }
});

let mailOptions = {
    from: config.username,
    to: 'mertens.achim@gmail.com',
    subject: 'Dies ist mein erster Test eine Mail aus Javascript',
    text: 'Hallo Achim, Wenn du das hier lesen kannst, bist du ein Genie ;-)'
};

transporter.sendMail(mailOptions, function(error, info){
    if (error) {
        console.log(error);
    } else {
        console.log('E-Mail erfolgreich gesendet: ' + info.response);
    }
});
