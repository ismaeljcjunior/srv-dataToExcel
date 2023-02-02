const express = require('express');
const axios = require('axios');
 
const app = express();
const port = 3000;

app.use(express.json());

app.post('/send-data', (req, res) => {
  const data = req.body;
  
  axios.post('http://localhost:3001/convert-to-excel', data)
    .then(response => {
      res.download(response.data);
    })
    .catch(error => {
      console.error(error);
      res.status(500).send('Error converting to Excel');
    });
});

app.listen(port, () => {
  console.log(`Server 1 listening at http://localhost:${port}`);
});
