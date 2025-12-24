import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import { ChevronDown, Search, Globe } from 'lucide-react';

const Header = () => {
  const [activeDropdown, setActiveDropdown] = useState(null);

  const menuItems = [
    {
      name: 'Services',
      items: [
        { name: 'Business consulting', link: '/services/business-consulting' },
        { name: 'Business process services', link: '/services/bps' },
        { name: 'Managed IT services', link: '/services/managed-it' },
        { name: 'Artificial intelligence', link: '/services/ai' },
        { name: 'Data analytics', link: '/services/data-analytics' },
        { name: 'Cloud & hybrid IT', link: '/services/cloud' },
        { name: 'View all services', link: '/services', featured: true }
      ]
    },
    {
      name: 'Industries',
      items: [
        { name: 'Banking', link: '/industries/banking' },
        { name: 'Government', link: '/industries/government' },
        { name: 'Health', link: '/industries/health' },
        { name: 'Manufacturing', link: '/industries/manufacturing' },
        { name: 'Retail', link: '/industries/retail' },
        { name: 'Energy & utilities', link: '/industries/energy' },
        { name: 'View all industries', link: '/industries', featured: true }
      ]
    },
    {
      name: 'Insights',
      items: [
        { name: 'Blog', link: '/blog' },
        { name: 'Case studies', link: '/case-studies' },
        { name: 'News', link: '/news' },
        { name: 'Events', link: '/events' }
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
            <svg width="80" height="40" viewBox="0 0 80 40" className="text-[#E31937]">
              <text x="10" y="28" fontSize="28" fontWeight="bold" fill="currentColor">CGI</text>
            </svg>
          </Link>

          {/* Navigation */}
          <nav className="hidden md:flex items-center space-x-8">
            {menuItems.map((item) => (
              <div
                key={item.name}
                className="relative"
                onMouseEnter={() => setActiveDropdown(item.name)}
                onMouseLeave={() => setActiveDropdown(null)}
              >
                {item.link ? (
                  <Link
                    to={item.link}
                    className="text-gray-800 hover:text-[#E31937] font-medium text-sm flex items-center transition-colors"
                  >
                    {item.name}
                  </Link>
                ) : (
                  <button className="text-gray-800 hover:text-[#E31937] font-medium text-sm flex items-center transition-colors">
                    {item.name}
                    <ChevronDown className="ml-1 h-4 w-4" />
                  </button>
                )}

                {/* Dropdown */}
                {item.items && activeDropdown === item.name && (
                  <div className="absolute left-0 top-full mt-2 w-64 bg-white shadow-lg border border-gray-200 py-2">
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
              to="/investors"
              className="text-gray-800 hover:text-[#E31937] font-medium text-sm transition-colors"
            >
              Investors
            </Link>
          </nav>

          {/* Right side actions */}
          <div className="flex items-center space-x-4">
            <Link
              to="/contact"
              className="text-gray-800 hover:text-[#E31937] font-medium text-sm transition-colors"
            >
              Contact
            </Link>
            <button className="p-2 text-gray-600 hover:text-[#E31937] transition-colors">
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
            </div>
          </div>
        </div>
      </div>
    </header>
  );
};

export default Header;
