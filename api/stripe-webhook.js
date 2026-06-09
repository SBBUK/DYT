const Stripe = require('stripe');

module.exports.config = { api: { bodyParser: false } };

const SCRIPT_URL = process.env.APPS_SCRIPT_URL;

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end();

  const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
  const sig = req.headers['stripe-signature'];

  let rawBody;
  try {
    rawBody = await new Promise((resolve, reject) => {
      const chunks = [];
      req.on('data', chunk => chunks.push(chunk));
      req.on('end', () => resolve(Buffer.concat(chunks)));
      req.on('error', reject);
    });
  } catch (err) {
    return res.status(400).send('Could not read body');
  }

  let event;
  try {
    event = stripe.webhooks.constructEvent(rawBody, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    console.error('Webhook signature error:', err.message);
    return res.status(400).send('Webhook error: ' + err.message);
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    const { name, handle, email, first, sessionId } = session.metadata;

    try {
      const url = `${SCRIPT_URL}?action=session_register&name=${encodeURIComponent(name)}&handle=${encodeURIComponent(handle)}&email=${encodeURIComponent(email)}&first=${first}&id=${encodeURIComponent(sessionId)}&paid=true`;
      await fetch(url);
    } catch (err) {
      console.error('Apps Script error:', err);
    }
  }

  res.status(200).json({ received: true });
};
