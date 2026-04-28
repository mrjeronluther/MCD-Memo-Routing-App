# MCD Memo Routing App

## Introduction
MCD Memo Routing App is a lightweight web application for creating, routing, and tracking internal memos. It is designed for MCD teams who need a simple, auditable way to submit memos, assign recipients, and follow routing status within the organization.

## Installation Instructions
Prerequisites:
- Node.js (v16+ recommended) and npm or yarn
- Git (to clone the repository)

Quick start:
1. Clone the repository:
   git clone https://github.com/mrjeronluther/MCD-Memo-Routing-App.git
2. Change into the project directory:
   cd MCD-Memo-Routing-App
3. Install dependencies:
   npm install
   or
   yarn install
4. Run the development server:
   npm start
   or
   yarn start
5. Open your browser at:
   http://localhost:3000
(If this is a static build, you can also run: npm run build && npx serve build)

## Usage Examples

Basic HTML form (example)
```html
<form id="memoForm">
  <input id="title" placeholder="Memo title" required />
  <textarea id="body" placeholder="Memo body" required></textarea>
  <input id="recipients" placeholder="Comma-separated recipients" />
  <button type="submit">Send Memo</button>
</form>
```

Example JavaScript usage (frontend)
```javascript
// Example: routeMemo is a small helper that posts a memo to the app backend
async function routeMemo(memo) {
  const res = await fetch('/api/memos', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(memo)
  });
  if (!res.ok) throw new Error('Failed to route memo');
  return res.json();
}

// Usage:
routeMemo({
  title: 'Weekly Ops Update',
  body: 'Please review attached targets and schedule.',
  recipients: ['Store Manager', 'Regional Ops'],
  priority: 'Normal'
})
  .then(response => console.log('Memo routed:', response))
  .catch(err => console.error(err));
```

CLI (local test)
```bash
# start dev server
npm start

# run tests (if available)
npm test
```

## Notes
- Adjust API endpoints and scripts according to your deployment/backend setup.
- Add authentication and access controls before deploying to production.

- For MCD Internal Use Only
