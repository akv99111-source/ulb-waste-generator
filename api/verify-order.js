export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  const { order_id } = req.body;

  if (!order_id) {
    return res.status(400).json({ message: 'Order ID is required' });
  }

  const CLIENT_ID = process.env.CASHFREE_APP_ID;
  const SECRET_KEY = process.env.CASHFREE_SECRET_KEY;
  const IS_PRODUCTION = process.env.CASHFREE_MODE === 'production';

  const baseUrl = IS_PRODUCTION 
    ? `https://api.cashfree.com/pg/orders/${order_id}`
    : `https://sandbox.cashfree.com/pg/orders/${order_id}`;

  try {
    const response = await fetch(baseUrl, {
      method: 'GET',
      headers: {
        'x-api-version': '2023-08-01',
        'x-client-id': CLIENT_ID,
        'x-client-secret': SECRET_KEY
      }
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.message || 'Error fetching order from Cashfree');
    }

    // Cashfree returns "PAID" if the transaction was successful
    if (data.order_status === 'PAID') {
      return res.status(200).json({ success: true, message: 'Payment verified' });
    } else {
      return res.status(400).json({ success: false, message: `Order status is ${data.order_status}. Payment not completed.` });
    }

  } catch (error) {
    return res.status(500).json({ success: false, message: error.message });
  }
}
