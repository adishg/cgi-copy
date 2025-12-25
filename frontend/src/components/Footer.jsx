import React from 'react';
import { Link } from 'react-router-dom';
import { Linkedin, Twitter, Facebook, Youtube } from 'lucide-react';

const Footer = () => {
  const footerSections = [
    {
      title: 'Services',
      links: [
        { name: 'Business consulting', url: '/services/business-consulting' },
        { name: 'Business process services', url: '/services/bps' },
        { name: 'Managed IT services', url: '/services/managed-it' },
        { name: 'Artificial intelligence', url: '/services/ai' },
        { name: 'Data analytics', url: '/services/data-analytics' },
        { name: 'Cloud & hybrid IT', url: '/services/cloud' }
      ]
    },
    {
      title: 'Industries',
      links: [
        { name: 'Banking', url: '/industries/banking' },
        { name: 'Government', url: '/industries/government' },
        { name: 'Health', url: '/industries/health' },
        { name: 'Manufacturing', url: '/industries/manufacturing' },
        { name: 'Retail', url: '/industries/retail' },
        { name: 'Energy & utilities', url: '/industries/energy' }
      ]
    },
    {
      title: 'Company',
      links: [
        { name: 'About us', url: '/about' },
        { name: 'Careers', url: '/careers' },
        { name: 'News', url: '/news' },
        { name: 'Investors', url: '/investors' },
        { name: 'Contact', url: '/contact' },
        { name: 'Locations', url: '/locations' }
      ]
    },
    {
      title: 'Insights',
      links: [
        { name: 'Blog', url: '/blog' },
        { name: 'Case studies', url: '/case-studies' },
        { name: 'Events', url: '/events' },
        { name: 'Voice of our clients', url: '/voice-of-clients' }
      ]
    }
  ];

  return (
    <footer className="bg-[#1A1A1A] text-white">
      <div className="max-w-7xl mx-auto px-6 py-12">
        {/* Main footer content */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8 mb-8">
          {footerSections.map((section) => (
            <div key={section.title}>
              <h3 className="text-lg font-semibold mb-4">{section.title}</h3>
              <ul className="space-y-2">
                {section.links.map((link) => (
                  <li key={link.name}>
                    <Link
                      to={link.url}
                      className="text-gray-400 hover:text-white text-sm transition-colors"
                    >
                      {link.name}
                    </Link>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>

        {/* Social media and bottom links */}
        <div className="border-t border-gray-800 pt-8">
          <div className="flex flex-col md:flex-row justify-between items-center">
            <div className="flex items-center space-x-6 mb-4 md:mb-0">
              <a href="https://linkedin.com" target="_blank" rel="noopener noreferrer" className="text-gray-400 hover:text-white transition-colors">
                <Linkedin className="h-5 w-5" />
              </a>
              <a href="https://twitter.com" target="_blank" rel="noopener noreferrer" className="text-gray-400 hover:text-white transition-colors">
                <Twitter className="h-5 w-5" />
              </a>
              <a href="https://facebook.com" target="_blank" rel="noopener noreferrer" className="text-gray-400 hover:text-white transition-colors">
                <Facebook className="h-5 w-5" />
              </a>
              <a href="https://youtube.com" target="_blank" rel="noopener noreferrer" className="text-gray-400 hover:text-white transition-colors">
                <Youtube className="h-5 w-5" />
              </a>
            </div>

            <div className="flex flex-wrap justify-center md:justify-end items-center space-x-6 text-sm text-gray-400">
              <Link to="/privacy" className="hover:text-white transition-colors">Privacy</Link>
              <Link to="/legal" className="hover:text-white transition-colors">Legal</Link>
              <Link to="/accessibility" className="hover:text-white transition-colors">Accessibility</Link>
              <Link to="/sitemap" className="hover:text-white transition-colors">Sitemap</Link>
            </div>
          </div>

          <div className="mt-6 text-center md:text-left text-sm text-gray-500">
            <p>&copy; 2025 CGI Inc. All rights reserved.</p>
          </div>
        </div>
      </div>
    </footer>
  );
};

export default Footer;
