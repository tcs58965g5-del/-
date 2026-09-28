import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import DevPanel from './DevPanel.tsx';
import './index.css';

const isDeveloperRoute = window.location.hash === '#developer';

const root = createRoot(document.getElementById('root')!);
root.render(
  <StrictMode>
    {isDeveloperRoute ? <DevPanel /> : <App />}
  </StrictMode>,
);

// Hide splash screen once React has painted
requestAnimationFrame(() => {
  requestAnimationFrame(() => {
    const splash = document.getElementById('injaz-splash');
    if (splash) {
      splash.classList.add('hide');
      setTimeout(() => splash.remove(), 600);
    }
  });
});
