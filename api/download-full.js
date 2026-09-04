// /api/download-full.js
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method not allowed' });

  const { order_id, configData } = req.body;

  // 1. Verify Payment Status with Cashfree API
  const cashfreeRes = await fetch(`https://api.cashfree.com/pg/orders/${order_id}`, {
    headers: {
      "x-api-version": "2023-08-01",
      "x-client-id": process.env.CASHFREE_APP_ID,
      "x-client-secret": process.env.CASHFREE_SECRET_KEY,
    }
  });

  const orderStatus = await cashfreeRes.json();

  if (orderStatus.order_status !== 'PAID') {
    return res.status(403).json({ success: false, message: 'Payment not verified.' });
  }

  // 2. Generate FULL 30/31-day dataset on server
  const fullMonthlyData = generateFullDataset(configData);

  // 3. Return full dataset or binary Excel stream
  return res.status(200).json({
    success: true,
    isPaid: true,
    data: fullMonthlyData
  });
}
