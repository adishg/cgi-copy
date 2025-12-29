import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import { news } from '../mockData';
import { Calendar, ArrowRight, FileText } from 'lucide-react';

const NewsPage = () => {
  const [selectedYear, setSelectedYear] = useState('2025');

  const years = ['2025', '2024', '2023'];

  // Filter news by year
  const filteredNews = news.filter(item => item.date.includes(selectedYear));

  // Additional news items for demonstration
  const allNews = [
    ...news,
    {
      id: 4,
      title: "Aventra Consulting announces strategic partnership with leading cloud provider",
      type: "Partnership announcement",
      date: "December 15, 2024",
      excerpt: "New partnership will enable CGI to deliver enhanced cloud migration and optimization services to enterprise clients globally.",
      image: "https://images.unsplash.com/photo-1559136555-9303baea8ebd?w=600&h=400&fit=crop",
      link: "/news/cloud-partnership"
    },
    {
      id: 5,
      title: "Aventra Consulting recognized as leader in IT services and consulting",
      type: "Industry recognition",
      date: "December 8, 2024",
      excerpt: "Independent analyst firm positions Aventra Consulting as a leader in global IT services market for fifth consecutive year.",
      image: "https://images.unsplash.com/photo-1552664730-d307ca884978?w=600&h=400&fit=crop",
      link: "/news/industry-recognition"
    },
    {
      id: 6,
      title: "Aventra Consulting launches AI-powered analytics platform for financial services",
      type: "Product launch",
      date: "November 28, 2024",
      excerpt: "New platform combines advanced AI capabilities with industry expertise to help financial institutions unlock data insights.",
      image: "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=600&h=400&fit=crop",
      link: "/news/ai-analytics-platform"
    }
  ];

  const newsToDisplay = selectedYear === '2025' 
    ? allNews.filter(item => item.date.includes('2025'))
    : allNews.filter(item => item.date.includes(selectedYear));

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-[#1A1A1A] text-white py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold mb-6">CGI News & Updates</h1>
          <p className="text-xl text-gray-300 max-w-3xl">
            Stay up to date with the latest news, announcements, and developments from CGI around the world.
          </p>
        </div>
      </section>

      {/* Year Filter */}
      <section className="bg-white border-b border-gray-200 sticky top-16 z-40">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center gap-4">
            <span className="text-gray-700 font-medium">Filter by year:</span>
            <div className="flex gap-2">
              {years.map((year) => (
                <button
                  key={year}
                  onClick={() => setSelectedYear(year)}
                  className={`px-4 py-2 rounded-md font-medium transition-colors ${
                    selectedYear === year
                      ? 'bg-[#E31937] text-white'
                      : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                  }`}
                >
                  {year}
                </button>
              ))}
            </div>
          </div>
        </div>
      </section>

      {/* Latest News - Featured */}
      {newsToDisplay.length > 0 && (
        <section className="py-12 bg-[#F8F8F8]">
          <div className="max-w-7xl mx-auto px-6">
            <h2 className="text-3xl font-bold text-gray-900 mb-8">Latest News</h2>
            <div className="bg-white rounded-lg overflow-hidden shadow-lg">
              <div className="grid md:grid-cols-5 gap-0">
                <div className="md:col-span-2 aspect-video md:aspect-auto">
                  <img
                    src={newsToDisplay[0].image}
                    alt={newsToDisplay[0].title}
                    className="w-full h-full object-cover"
                  />
                </div>
                <div className="md:col-span-3 p-8 md:p-12 flex flex-col justify-center">
                  <div className="flex items-center gap-4 mb-4">
                    <span className="px-3 py-1 bg-[#E31937] text-white text-sm font-semibold rounded uppercase">
                      {newsToDisplay[0].type}
                    </span>
                    <span className="flex items-center gap-2 text-sm text-gray-600">
                      <Calendar className="h-4 w-4" />
                      {newsToDisplay[0].date}
                    </span>
                  </div>
                  <h2 className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
                    {newsToDisplay[0].title}
                  </h2>
                  <p className="text-lg text-gray-700 mb-6">
                    {newsToDisplay[0].excerpt}
                  </p>
                  <Link
                    to={newsToDisplay[0].link}
                    className="inline-flex items-center text-[#E31937] font-semibold hover:underline"
                  >
                    Read full story
                    <ArrowRight className="ml-2 h-5 w-5" />
                  </Link>
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {/* All News Grid */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-3xl font-bold text-gray-900 mb-8">
            All News - {selectedYear}
          </h2>
          <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
            {newsToDisplay.map((item) => (
              <Link
                key={item.id}
                to={item.link}
                className="group bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-shadow"
              >
                <div className="aspect-video overflow-hidden">
                  <img
                    src={item.image}
                    alt={item.title}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <div className="p-6">
                  <div className="flex items-center justify-between mb-3">
                    <span className="inline-flex items-center gap-1 px-2 py-1 bg-gray-100 text-gray-700 text-xs font-medium rounded uppercase">
                      <FileText className="h-3 w-3" />
                      {item.type}
                    </span>
                    <span className="flex items-center gap-1 text-sm text-gray-600">
                      <Calendar className="h-4 w-4" />
                      {item.date}
                    </span>
                  </div>
                  <h3 className="text-xl font-bold text-gray-900 mb-3 group-hover:text-[#E31937] transition-colors">
                    {item.title}
                  </h3>
                  {item.excerpt && (
                    <p className="text-gray-700 mb-4 line-clamp-2">
                      {item.excerpt}
                    </p>
                  )}
                  <span className="inline-flex items-center text-[#E31937] font-semibold group-hover:underline">
                    Read more
                    <ArrowRight className="ml-2 h-4 w-4 group-hover:translate-x-1 transition-transform" />
                  </span>
                </div>
              </Link>
            ))}
          </div>
        </div>
      </section>

      {/* Media Contact Section */}
      <section className="py-16 bg-[#F8F8F8]">
        <div className="max-w-4xl mx-auto px-6">
          <div className="bg-white border border-gray-200 p-8 md:p-12">
            <h2 className="text-3xl font-bold text-gray-900 mb-4">Media inquiries</h2>
            <p className="text-lg text-gray-700 mb-6">
              For media inquiries, interview requests, or press information, please contact our corporate communications team.
            </p>
            <div className="space-y-3 text-gray-700">
              <p>
                <strong>Email:</strong>{' '}
                <a href="mailto:media@cgi.com" className="text-[#E31937] hover:underline">
                  media@cgi.com
                </a>
              </p>
              <p>
                <strong>Phone:</strong> +1 514 841 3355
              </p>
            </div>
            <div className="mt-8">
              <Link
                to="/contact"
                className="inline-block bg-[#E31937] text-white px-6 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
              >
                Contact us
              </Link>
            </div>
          </div>
        </div>
      </section>

      {/* Subscribe Section */}
      <section className="py-16 bg-white">
        <div className="max-w-4xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold text-gray-900 mb-4">
            Stay informed
          </h2>
          <p className="text-lg text-gray-700 mb-8">
            Subscribe to receive the latest CGI news and announcements directly to your inbox.
          </p>
          <form className="flex flex-col sm:flex-row gap-4 max-w-xl mx-auto">
            <input
              type="email"
              placeholder="Enter your email"
              className="flex-1 px-4 py-3 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-[#E31937]"
            />
            <button
              type="submit"
              className="px-8 py-3 bg-[#E31937] text-white font-semibold rounded-md hover:bg-[#C01530] transition-colors whitespace-nowrap"
            >
              Subscribe
            </button>
          </form>
        </div>
      </section>
    </div>
  );
};

export default NewsPage;
