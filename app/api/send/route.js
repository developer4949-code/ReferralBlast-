import nodemailer from 'nodemailer';

export async function POST(request) {
  try {
    const { senderEmail, appPassword, to, cc, subject, body, attachments } = await request.json();

    if (!senderEmail || !appPassword || !to) {
      return Response.json({ error: 'Missing required fields' }, { status: 400 });
    }

    // Configure Nodemailer transporter
    const transporter = nodemailer.createTransport({
      host: 'smtp.gmail.com',
      port: 465,
      secure: true, // Use SSL
      auth: {
        user: senderEmail,
        pass: appPassword,
      },
    });

    // Format attachments for Nodemailer
    const mailAttachments = attachments.map(att => ({
      filename: att.filename,
      content: att.content,
      encoding: 'base64',
      contentType: att.contentType
    }));

    // Send Mail
    const mailOptions = {
      from: senderEmail,
      to,
      cc: cc || undefined,
      subject,
      text: body, // Plain text body
      html: `<div style="font-family: serif, 'Times New Roman', Times; font-size: 16px;">${body.replace(/\n/g, '<br>')}</div>`, // HTML body
      attachments: mailAttachments
    };

    const info = await transporter.sendMail(mailOptions);

    return Response.json({ success: true, messageId: info.messageId });
  } catch (error) {
    console.error("Mail Send Error:", error);
    return Response.json({ error: error.message }, { status: 500 });
  }
}
