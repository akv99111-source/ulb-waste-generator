export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method not allowed' });

  try {
    const { amount, customerName, customerPhone } = req.body;

    const response = await fetch("https://api.cashfree.com/pg/orders", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-version": "2023-08-01",
        "x-client-id": process.env.CASHFREE_APP_ID,
        "x-client-secret": process.env.CASHFREE_SECRET_KEY,
      },
      body: JSON.stringify({
        order_amount: amount,
        order_currency: "INR",
        customer_details: {
          customer_id: `cust_${Date.now()}`,
          customer_name: customerName || "Customer",
          customer_phone: customerPhone || "9876543210", // Pass user's real 10-digit number
        },
      }),
    });

    const data = await response.json();
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ message: error.message });
  }
}
