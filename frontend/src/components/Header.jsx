import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import { ChevronDown, Search, Globe, Menu, X } from 'lucide-react';

const Header = () => {
  const [activeDropdown, setActiveDropdown] = useState(null);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [mobileSubmenuOpen, setMobileSubmenuOpen] = useState(null);

  const menuItems = [
    {
      name: 'Home',
      link: '/'
    },
    {
      name: 'About',
      link: '/about'
    },
    {
      name: 'Services',
      items: [
        { name: 'Business consulting', link: '/services' },
        { name: 'Business process services', link: '/services' },
        { name: 'Managed IT services', link: '/services' },
        { name: 'Artificial intelligence', link: '/services' },
        { name: 'Data analytics', link: '/services' },
        { name: 'Cloud & hybrid IT', link: '/services' },
        { name: 'View all services', link: '/services', featured: true }
      ]
    },
    {
      name: 'Industries',
      items: [
        { name: 'Banking', link: '/industries' },
        { name: 'Government', link: '/industries' },
        { name: 'Health', link: '/industries' },
        { name: 'Manufacturing', link: '/industries' },
        { name: 'Retail', link: '/industries' },
        { name: 'Energy & utilities', link: '/industries' },
        { name: 'View all industries', link: '/industries', featured: true }
      ]
    },
    {
      name: 'Insights',
      items: [
        { name: 'Blog', link: '/blog' },
        // { name: 'Case studies', link: '/case-studies' },
        { name: 'News', link: '/news' },
        // { name: 'Events', link: '/events' }
      ]
    },
    {
      name: 'Careers',
      link: '/careers'
    }
  ];

  return (
    <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
      <div className="max-w-7xl mx-auto px-6">
        <div className="flex items-center justify-between h-16">
          {/* Logo */}
          <Link to="/" className="flex-shrink-0">
            <img style={{height:"60%"}} src="/logo.PNG" alt="Company Logo" className="h-8 w-auto" />
          </Link>

          {/* Desktop Navigation */}
          <nav className="hidden md:flex items-center space-x-8">
            {menuItems.map((item) => (
              <div
                key={item.name}
                className="relative group"
                onMouseEnter={() => setActiveDropdown(item.name)}
                onMouseLeave={() => setActiveDropdown(null)}
              >
                {item.link ? (
                  <Link
                    to={item.link}
                    className="text-gray-800 hover:text-[#E31937] font-medium text-sm flex items-center transition-colors py-6"
                  >
                    {item.name}
                  </Link>
                ) : (
                  <button className="text-gray-800 hover:text-[#E31937] font-medium text-sm flex items-center transition-colors py-6">
                    {item.name}
                    <ChevronDown className="ml-1 h-4 w-4" />
                  </button>
                )}

                {/* Dropdown */}
                {item.items && activeDropdown === item.name && (
                  <div className="absolute left-0 top-full w-64 bg-white shadow-lg border border-gray-200 py-2 z-50">
                    {item.items.map((subItem) => (
                      <Link
                        key={subItem.name}
                        to={subItem.link}
                        className={`block px-4 py-2 text-sm hover:bg-gray-50 transition-colors ${
                          subItem.featured
                            ? 'text-[#E31937] font-semibold border-t border-gray-200 mt-2 pt-3'
                            : 'text-gray-700 hover:text-[#E31937]'
                        }`}
                      >
                        {subItem.name}
                      </Link>
                    ))}
                  </div>
                )}
              </div>
            ))}

            <Link
              to="/team"
              className="text-gray-800 hover:text-[#E31937] font-medium text-sm transition-colors"
            >
              Team
            </Link>
          </nav>

          {/* Right side actions */}
          <div className="hidden md:flex items-center space-x-4">
            <Link
              to="/contact"
              className="text-gray-800 hover:text-[#E31937] font-medium text-sm transition-colors"
            >
              Contact
            </Link>
            {/* <button className="p-2 text-gray-600 hover:text-[#E31937] transition-colors">
              <Search className="h-5 w-5" />
            </button>
            <button className="flex items-center space-x-1 text-gray-800 hover:text-[#E31937] transition-colors">
              <Globe className="h-5 w-5" />
              <span className="text-sm font-medium">Global</span>
            </button>
            <div className="flex items-center space-x-1">
              <button className="text-sm font-medium text-gray-800 hover:text-[#E31937] transition-colors">
                EN
              </button>
              <ChevronDown className="h-4 w-4 text-gray-600" />
            </div> */}
          </div>

          {/* Mobile menu button */}
          <button
            className="md:hidden p-2 text-gray-800 hover:text-[#E31937] transition-colors"
            onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            aria-label="Toggle menu"
          >
            {mobileMenuOpen ? <X className="h-6 w-6" /> : <Menu className="h-6 w-6" />}
          </button>
        </div>
      </div>

      {/* Mobile Menu */}
      {mobileMenuOpen && (
        <div className="md:hidden bg-white border-t border-gray-200 shadow-lg max-h-screen overflow-y-auto">
          <div className="px-6 py-4 space-y-4">
            {menuItems.map((item) => (
              <div key={item.name}>
                {item.link ? (
                  <Link
                    to={item.link}
                    className="block py-2 text-gray-800 hover:text-[#E31937] font-medium transition-colors"
                    onClick={() => setMobileMenuOpen(false)}
                  >
                    {item.name}
                  </Link>
                ) : (
                  <>
                    <button
                      className="w-full flex items-center justify-between py-2 text-gray-800 hover:text-[#E31937] font-medium transition-colors"
                      onClick={() => setMobileSubmenuOpen(mobileSubmenuOpen === item.name ? null : item.name)}
                    >
                      {item.name}
                      <ChevronDown
                        className={`h-4 w-4 transition-transform ${
                          mobileSubmenuOpen === item.name ? 'rotate-180' : ''
                        }`}
                      />
                    </button>
                    {mobileSubmenuOpen === item.name && item.items && (
                      <div className="pl-4 space-y-2 mt-2">
                        {item.items.map((subItem) => (
                          <Link
                            key={subItem.name}
                            to={subItem.link}
                            className={`block py-2 text-sm transition-colors ${
                              subItem.featured
                                ? 'text-[#E31937] font-semibold'
                                : 'text-gray-700 hover:text-[#E31937]'
                            }`}
                            onClick={() => {
                              setMobileMenuOpen(false);
                              setMobileSubmenuOpen(null);
                            }}
                          >
                            {subItem.name}
                          </Link>
                        ))}
                      </div>
                    )}
                  </>
                )}
              </div>
            ))}

            <Link
              to="/investors"
              className="block py-2 text-gray-800 hover:text-[#E31937] font-medium transition-colors"
              onClick={() => setMobileMenuOpen(false)}
            >
              Team
            </Link>

            <div className="border-t border-gray-200 pt-4 space-y-3">
              <Link
                to="/contact"
                className="block py-2 text-gray-800 hover:text-[#E31937] font-medium transition-colors"
                onClick={() => setMobileMenuOpen(false)}
              >
                Contact
              </Link>
              {/* <button className="flex items-center space-x-2 py-2 text-gray-800 hover:text-[#E31937] transition-colors">
                <Globe className="h-5 w-5" />
                <span className="font-medium">Global</span>
              </button>
              <button className="flex items-center space-x-2 py-2 text-gray-800 hover:text-[#E31937] transition-colors">
                <span className="font-medium">EN</span>
                <ChevronDown className="h-4 w-4" />
              </button> */}
            </div>
          </div>
        </div>
      )}
    </header>
  );
};

export default Header;
