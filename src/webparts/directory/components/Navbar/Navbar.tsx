import * as React from 'react';
import { useState } from 'react';
import { Link } from 'react-router-dom'; // Import Link from React Router
import { BrowserRouter as Router } from 'react-router-dom'; // Import BrowserRouter
import styles from './Navbar.module.scss';

const Navbar: React.FC = () => {
  const [isActive, setIsActive] = useState(false);

  const toggleActiveClass = () => {
    setIsActive(!isActive);
  };

  const removeActive = () => {
    setIsActive(false);
  };

  return (
    <Router> {/* Wrap your Navbar component with Router */}
      <div className="App">
        <header className="App-header">
          <nav className={`${styles.navbar}`}>
            <Link to="/CnexiaForEveryone" className={`${styles.logo}`}>
              {/* Your SVG Logo */}
            </Link>
            <ul className={`${styles.navMenu} ${isActive ? styles.active : ''}`}>
              <li onClick={removeActive}>
                <Link to="/CnexiaForEveryone/people-culture" className={`${styles.navLink}`}>
                  People & Culture
                </Link>
              </li>
              <li onClick={removeActive}>
                <Link to="/CnexiaForEveryone/career" className={`${styles.navLink}`}>
                  Career
                </Link>
              </li>
              <li onClick={removeActive}>
                <Link to="/CnexiaForEveryone/communication" className={`${styles.navLink}`}>
                  Communication
                </Link>
              </li>
              <li onClick={removeActive}>
                <Link to="/CnexiaForEveryone/support" className={`${styles.navLink}`}>
                  Support
                </Link>
              </li>
            </ul>
            <div
              className={`${styles.hamburger} ${isActive ? styles.active : ''}`}
              onClick={toggleActiveClass}
            >
              {/* Your hamburger icon */}
            </div>
          </nav>
        </header>
      </div>
    </Router>
  );
};

export default Navbar;
