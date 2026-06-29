const Clerk = require('@clerk/backend');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', 'https://anttro.com');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const authHeader = req.headers['authorization'];
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'No session token provided.' });
  }

  const token = authHeader.replace('Bearer ', '');

  try {
    const clerkClient = Clerk.createClerkClient({
      secretKey: process.env.CLERK_SECRET_KEY,
    });

    // Verify the token
    const payload = await clerkClient.verifyToken(token);
    const userId = payload.sub;

    // Get the user's metadata to find their team ID
    const user = await clerkClient.users.getUser(userId);
    const teamId = user.publicMetadata?.teamId || user.privateMetadata?.teamId || null;

    if (!teamId) {
      return res.status(403).json({ error: 'No team assigned to this account. Contact Anttro support.' });
    }

    return res.status(200).json({ teamId });

  } catch (err) {
    console.error('Clerk verify error:', err);
    return res.status(401).json({ error: 'Invalid or expired session. Please sign in again.' });
  }
};
