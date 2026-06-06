const Stripe = require('stripe');

module.exports.config = { api: { bodyParser: false } };

const SCRIPT_URL = process.env.APPS_SCRIPT_URL;

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end();

  const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
  const sig = req.headers['stripe-signature'];

  let event;
  try {
    event = stripe.webhooks.constructEvent(req.body, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    console.error('Webhook signature error:', err.message);
    return res.status(400).send('Webhook error: ' + err.message);
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    const { name, handle, email, first, sessionId, sessionName, sessionDate, sessionTime } = session.metadata;

    // Register the spot in Apps Script now payment is confirmed
    try {
      const url = `${SCRIPT_URL}?action=session_register&name=${encodeURIComponent(name)}&handle=${encodeURIComponent(handle)}&email=${encodeURIComponent(email)}&first=${first}&id=${encodeURIComponent(sessionId)}&paid=true`;
      await fetch(url);
    } catch (err) {
      console.error('Apps Script error:', err);
    }
  }

  res.status(200).json({ received: true });
};
