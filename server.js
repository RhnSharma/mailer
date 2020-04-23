require('dotenv').config();
const CronJob = require('cron').CronJob;
const mongoose = require('mongoose');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const Schema = mongoose.Schema;
const SubmissionSchema = new Schema({},{strict : false});
const Submission = mongoose.model('Submission', SubmissionSchema);
let newWB, newWS;

mongoose.connect(process.env.MONGO_URI, {useNewUrlParser : true, useUnifiedTopology : true })
.then(async () => {
  console.log('Connected to database.')
})
.catch(err => {
  console.log(err)
});

let transporter = nodemailer.createTransport({
  service : 'gmail',
  auth : {
    user : process.env.EMAIL,
    pass : process.env.PASSWORD
  }
});

let job = new CronJob('0 * * * *', async function() {
  const submissions = await Submission.find({ newDoc : true });
  if(submissions.length > 0){
    console.log(`${submissions.length} new submission(s)!`);
    const updatedSubmissions = submissions.map(submission => {
      return {
        name : submission._doc.name,
        email : submission._doc.email,
        message : submission._doc.message,
        time : submission._doc.createdAt
      }
    })
    newWB = xlsx.utils.book_new();
    newWS = xlsx.utils.json_to_sheet(updatedSubmissions);
    xlsx.utils.book_append_sheet(newWB, newWS, "Submission Data");
    xlsx.writeFile(newWB, 'newsubmission(s).xlsx');
    let mailOptions = {
      from : process.env.EMAIL,
      to : 'rhnsharma5113@gmail.com',
      subject : 'New Submission(s)!',
      text : 'New submission(s)!',
      attachments : [
        { filename : 'newsubmission(s).xlsx', path : './newsubmission(s).xlsx' }
      ]
    };
    transporter.sendMail(mailOptions)
    .then(res => {
      console.log('Email sent!!');
    })
    .catch(err => {
      console.log('Error occurs',err);
    });
    Submission.updateMany(
      { newDoc : true },
      { $set: { newDoc: false } },
      function(err, result) {
        if (err) {
          console.log(err);
        } else {
          console.log(result);
        }
      }
    );
  } else {
    console.log('No new submission!');
  }
}, null, true, 'Asia/Kolkata');
job.start();