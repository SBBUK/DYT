const Stripe = require('stripe');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const stripe = Stripe(process.env.STRIPE_SECRET_KEY);

  const { name, handle, email, first, sessionId, sessionName, sessionDate, sessionTime, sessionLocation, amount, promoCode } = req.body;

  if (!name || !handle || !email || !sessionId) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  // Server-side promo validation — SBBMEMBER locks price to £6 (600 pence)
  let finalAmount = amount || 1000;
  if ((promoCode || '').toUpperCase() === 'SBBMEMBER' && finalAmount > 600) {
    finalAmount = 600;
  }

  try {
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [{
        price_data: {
          currency: 'gbp',
          product_data: {
            name: sessionName || 'DYT Session',
            description: sessionDate + ' · ' + sessionTime + (sessionLocation ? ' · ' + sessionLocation : ''),
          },
          unit_amount: finalAmount,
        },
        quantity: 1,
      }],
      mode: 'payment',
      customer_email: email,
      metadata: {
        name,
        handle,
        email,
        first: first ? 'true' : 'false',
        sessionId,
        sessionName: sessionName || '',
        sessionDate: sessionDate || '',
        sessionTime: sessionTime || '',
      },
      success_url: process.env.SITE_URL + '/session.html?id=' + sessionId + '&payment=success&handle=' + encodeURIComponent(handle),
      cancel_url: process.env.SITE_URL + '/session.html?id=' + sessionId + '&payment=cancelled',
    });

    return res.status(200).json({ url: session.url });
  } catch (err) {
    console.error('Stripe error:', err);
    return res.status(500).json({ error: err.message });
  }
};
