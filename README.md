# FMAERO Smart Tracking System

FMAERO Smart Tracking System is a web-based inventory tracking project for warehouse and material operations. It includes role-based access for admin, manager, storekeeper, and site supervisor users, together with Firebase Hosting and Firebase Realtime Database integration.

## Live Links

- Live website: [https://fmaero-smart-tracking-system.web.app](https://fmaero-smart-tracking-system.web.app)
- GitHub repository: [https://github.com/Kecik2309/fmaero-smart-tracking-system](https://github.com/Kecik2309/fmaero-smart-tracking-system)

## Features

- Login system with role-based access
- Admin dashboard and worker management
- Manager, storekeeper, and supervisor dashboards
- Inventory and material tracking workflow
- Firebase Hosting deployment
- Firebase Realtime Database integration

## Project Structure

- `login.html` - Login page
- `index.html` - Main inventory dashboard
- `admin-dashboard.html` - Admin dashboard
- `admin-workers.html` - Worker management page
- `manager-dashboard.html` - Manager dashboard
- `storekeeper-dashboard.html` - Storekeeper dashboard
- `supervisor-dashboard.html` - Site supervisor dashboard
- `style.css` - Main shared styling
- `app.js` - Dashboard logic

## Pages

- `login.html` - Main entry page
- `forgot-password.html` - Password recovery page
- `index.html` - Inventory operations page after login

## How To Use

1. Open the live website.
2. Sign in with your configured user account.
3. Access the dashboard based on your assigned role.
4. Manage inventory, workers, and tracking workflow from the related pages.

## Testing Notes

- Check that the root URL redirects to `login.html`
- Test login flow for each role
- Confirm dashboard navigation works correctly
- Verify Firebase Realtime Database reads and writes are working
- Re-deploy after any update to the HTML, CSS, or JavaScript files

## FYP Project Note

If this repository is used for your Final Year Project, you can later add:

- University name
- Faculty or programme
- Supervisor name
- Project submission semester

## Deployment

Use the following commands in the project folder:

```cmd
git add .
git commit -m "update project"
git push
firebase.cmd deploy --only hosting
```

## Tech Stack

- HTML
- CSS
- JavaScript
- Firebase Hosting
- Firebase Realtime Database

## Author

- NUR NAJMINA
