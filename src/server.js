require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');

const app = express();
const port = process.env.PORT || 3000;
const clientState = process.env.CLIENT_STATE;
const graphToken = process.env.GRAPH_TOKEN; // Add this to your .env file

app.use(bodyParser.json());

// Webhook endpoint
app.post('/webhook', async (req, res) => {
  // Validation token for initial handshake
  if (req.query && req.query.validationToken) {
    console.log('Validation token received:', req.query.validationToken);
    return res.status(200).send(req.query.validationToken);
  }

  const notifications = req.body.value;

  for (const notification of notifications) {
    if (notification.clientState !== clientState) {
      console.warn('Invalid clientState received:', notification.clientState);
      continue;
    }

    console.log('🔔 Webhook Notification Received:');
    console.log(JSON.stringify(notification, null, 2));

    const resource = notification.resource; // e.g., drives/{drive-id}/items/{item-id}

    if (graphToken) {
      try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/${resource}`, {
          headers: {
            Authorization: `Bearer ${graphToken}`
          }
        });
        console.log('📄 File Metadata:', response.data);
      } catch (error) {
        console.error('❌ Error fetching file metadata:', error.response?.data || error.message);
      }
    } else {
      console.log('⚠️ Skipping metadata fetch — GRAPH_TOKEN not set.');
    }
  }

  res.sendStatus(202);
});

app.listen(port, () => {
  console.log(`✅ Webhook listener running on port ${port}`);
});