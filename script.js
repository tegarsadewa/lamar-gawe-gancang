const fs = require("fs");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const { exec } = require("child_process");
const nodemailer = require("nodemailer");
const path = require("path");

const cariTanggal = () => {
  const angkaTanggal = new Date().getDate();
  const cariNamaBulan = () => {
    const bulan = new Date().getMonth() + 1;
    if (bulan == 1) return "Januari";
    if (bulan == 2) return "Februari";
    if (bulan == 3) return "Maret";
    if (bulan == 4) return "April";
    if (bulan == 5) return "Mei";
    if (bulan == 6) return "Juni";
    if (bulan == 7) return "Juli";
    if (bulan == 8) return "Agustus";
    if (bulan == 9) return "September";
    if (bulan == 10) return "Oktober";
    if (bulan == 11) return "November";
    if (bulan == 12) return "Desember";
  };
  const namaBulan = cariNamaBulan();
  const tahun = new Date().getFullYear();
  const tanggal = `${angkaTanggal} ${namaBulan} ${tahun}`;
  return tanggal;
};

// FORM EMAIL 👇
const data = {
  email: "emailperusahaan@gmail.com", // 👈 Email perusahaan yang dituju
  subject: "Hokage_Uzumaki Naruto", // 👈 Subjek email
  perusahaan: "PT. Mencari Cinta Sejati", // 👈 Nama perusahaan
  posisi: "Hokage", // 👈 Posisi yang dilamar
  tanggal: cariTanggal(),
};

const buatWord = (data) => {
  const outputDir = path.resolve(__dirname, "output");
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const content = fs.readFileSync("template.docx", "binary");

  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    delimiters: { start: "{{", end: "}}" },
  });

  try {
    doc.render(data);
  } catch (error) {
    console.error("Error saat rendering dokumen:", error);
    return;
  }

  const buffer = doc.getZip().generate({ type: "nodebuffer" });

  const outputPath = `output/Naruto Uzumaki - ${data.posisi}.docx`;
  fs.writeFileSync(outputPath, buffer);

  console.log(`Dokumen berhasil dibuat: ${outputPath}`);
  console.log(" ￣へ￣ ");
};

const buatPdf = async (data) => {
  const pathLibreOffice = `"C:/Program Files/LibreOffice/program/soffice.exe"`;

  // Sesuaikan direktori folder kalian 👇
  const file = `C:/Users/Naruto Uzumaki/Desktop/YTTA/test/output/Naruto Uzumaki - ${data.posisi}`;

  const command = `${pathLibreOffice} --headless --convert-to pdf "${file}".docx --outdir "./output"`;

  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error("Convert PDF Gagal:", error.message);
      return;
    }

    if (fs.existsSync(`${file}.pdf`)) {
      console.log("Convert PDF Berhasil:", `${file}.pdf`);
      console.log("（￣︶￣）↗　");
      kirimEmail(data);
    } else {
      console.log("Convert PDF Gagal:", `${file}.pdf`);
      console.log("(┬┬﹏┬┬)");
    }
  });
};

const kirimEmail = (data) => {
  let transporter = nodemailer.createTransport({
    service: "gmail",
    port: 465,
    secure: true,
    logger: true,
    debug: false,
    secureConnection: false,
    auth: {
      user: "narutouzumaki@gmail.com", // 👈 Ganti jadi email Lamar Gawe
      pass: "wkwkwkwkwkwkw", // 👈 Ganti jadi password email Lamar Gawe
    },
    tls: {
      rejectUnauthorized: true,
    },
  });

  let mailOptions = {
    from: "narutouzumaki@gmail.com", // 👈 Ganti jadi email Lamar Gawe
    to: `${data.email}`,
    subject: `${data.subject}`,
    // Ganti Pesan Email 👇
    text: `Dear Bapak/Ibu HRD,\nPerkenalkan nama saya Naruto Uzumaki, lulusan dari Universitas Sriwijaya dengan jurusan D3 Manajemen Informatika.\n\nSaya memiliki minat yang besar untuk dapat bekerja diposisi ${data.posisi}. Karena selama saya kuliah, saya mengamati dengan baik bahwa untuk bekerja diposisi ini membutuhkan keahlian komputer. Saya mempelajari hal tersebut dengan baik selama saya kuliah sehingga skill yang saya miliki dapat berguna untuk posisi yg Bapak/Ibu butuhkan.\n\nJika Bapak/Ibu segera membutuhkan posisi ini, panggil saja saya karena saya siap untuk melakukan proses seleksi kapanpun. Berikut kontak saya yang bisa dihubungi, dan terlampir berkas lampiran saya seperti CV, KTP, SKCK, ijazah dan transkrip nilai.\n\nTerimakasih sudah bersedia membaca surat lamaran saya, have a good day!\n\nThanks & Regards,\nNaruto Uzumaki\n08969696969`,
    attachments: [
      {
        filename: `Naruto Uzumaki - ${data.posisi}.pdf`,
        path: `./output/Naruto Uzumaki - ${data.posisi}.pdf`,
      },
    ],
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log("(┬┬﹏┬┬)\nError occurred:", error);
    } else {
      console.log("Email Terkirim:", info.response);
      console.log(`Email: ${data.email}`);
      console.log(`Subject: ${data.subject}`);
      console.log(`Perusahaan: ${data.perusahaan}`);
      console.log(`Posisi: ${data.posisi}`);
      console.log("(～￣▽￣)～");
    }
  });
};

const processAndSendEmail = async (data) => {
  buatWord(data);
  await buatPdf(data);
};

processAndSendEmail(data);
