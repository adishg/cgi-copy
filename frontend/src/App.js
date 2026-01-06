import React, { useEffect } from 'react';
import './App.css';
import { BrowserRouter, Routes, Route, useLocation } from 'react-router-dom';
import Header from './components/Header';
import Footer from './components/Footer';
import Home from './pages/Home';
import ServicesPage from './pages/ServicesPage';
import IndustriesPage from './pages/IndustriesPage';
import CareersPage from './pages/CareersPage';
import ContactPage from './pages/ContactPage';
import AboutPage from './pages/AboutPage';
import ExelenceDeliveredPage from './pages/ExelenceDeliveredPage';
import { FloatingWhatsApp } from 'react-floating-whatsapp'
import BlogPage from './pages/BlogPage';
import NewsPage from './pages/NewsPage';
import TeamPage from './pages/TeamPage';
import BlogDetailPage from './pages/BlogDetailPage';
import NewsDetailPage from './pages/NewsDetailPage';

// Component to handle scrolling to hash sections
const ScrollToHash = () => {
  const location = useLocation();
  
  useEffect(() => {
    if (location.hash) {
      // Give time for the page to render before scrolling
      const timeoutId = setTimeout(() => {
        const elementId = location.hash.substring(1); // Remove '#' from hash
        const element = document.getElementById(elementId);
        if (element) {
          // Get header height for offset (sticky header)
          const headerHeight = document.querySelector('header')?.offsetHeight || 80;
          const elementPosition = element.getBoundingClientRect().top + window.pageYOffset;
          const offsetPosition = elementPosition - headerHeight - 20; // 20px extra padding
          
          window.scrollTo({
            top: offsetPosition,
            behavior: 'smooth'
          });
        }
      }, 100); // Delay to ensure page has rendered
      
      return () => clearTimeout(timeoutId);
    } else {
      // Scroll to top when navigating to a page without hash
      window.scrollTo(0, 0);
    }
  }, [location]);
  
  return null;
};

function App() {
  return (
    <div className="App">
      <BrowserRouter>
        <Header />
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/services" element={<ServicesPage />} />
          <Route path="/services/:serviceId" element={<ServicesPage />} />
          <Route path="/industries" element={<IndustriesPage />} />
          <Route path="/industries/:industryId" element={<IndustriesPage />} />
          <Route path="/exelence-delivered" element={<ExelenceDeliveredPage />} />
          {/* <Route path="/casestudies" element={<CaseStudiesPage />} /> */}
          <Route path="/exelence-delivered/:caseId" element={<ExelenceDeliveredPage />} />
          <Route path="/careers" element={<CareersPage />} />
          <Route path="/contact" element={<ContactPage />} />
          <Route path="/about" element={<AboutPage />} />
          <Route path="/team" element={<TeamPage />} />
          <Route path="/blog" element={<BlogPage />} />
           <Route path="/blog/:blogId" element={<BlogDetailPage />} />
          <Route path="/news" element={<NewsPage />} />
           <Route path="/news/:newsId" element={<NewsDetailPage />} />
          <Route path="/team" element={<TeamPage />} />
          <Route path="*" element={<Home />} />
        </Routes>
        <Footer />
      </BrowserRouter>
      <FloatingWhatsApp 
        phoneNumber="919871833119" // Include country code
        accountName="Support"
        chatMessage="Hello! How can we help?"
        avatar="https://gravatar.com/avatar/6bed0a553547181a561564cc2daf8583?s=400&d=robohash&r=x"
        statusMessage="Typically replies within a day"
        notification
        notificationSound
      />
    </div>
  );
}

export default App;
