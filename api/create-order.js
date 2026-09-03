export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  const { amount, customerName } = req.body;

  const CLIENT_ID = process.env.CASHFREE_APP_ID;
  const SECRET_KEY = process.env.CASHFREE_SECRET_KEY;
  const IS_PRODUCTION = process.env.CASHFREE_MODE === 'production';

  const baseUrl = IS_PRODUCTION 
    ? 'https://api.cashfree.com/pg/orders'
    : 'https://sandbox.cashfree.com/pg/orders';

  const orderPayload = {
    order_id: `SWM_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
    order_amount: Number(amount),
    order_currency: 'INR',
    customer_details: {
      customer_id: `CUST_${Date.now()}`,
      customer_name: customerName || 'ULB User',
      customer_email: 'user@example.com',
      customer_phone: '9999999999'
    }
  };

  try {
    const response = await fetch(baseUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-version': '2023-08-01',
        'x-client-id': CLIENT_ID,
        'x-client-secret': SECRET_KEY
      },
      body: JSON.stringify(orderPayload)
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.message || 'Error creating Cashfree order');
    }

    return res.status(200).json({ 
      payment_session_id: data.payment_session_id, 
      order_id: data.order_id 
    });
  } catch (error) {
    return res.status(500).json({ message: error.message });
  }
}
