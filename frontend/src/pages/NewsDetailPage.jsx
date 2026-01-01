import React from 'react';
import { Link, useParams } from 'react-router-dom';
import { news } from '../mockData';
import { Calendar, ArrowLeft, Share2, Linkedin, Twitter, Facebook, FileText } from 'lucide-react';

const NewsDetailPage = () => {
  const { newsId } = useParams();
  
  // Find the news item by ID or use first item as default
  const newsItem = news.find(n => n.link.includes(newsId)) || news[0];
  
  // Get related news (exclude current item)
  const relatedNews = news.filter(n => n.id !== newsItem.id).slice(0, 3);

  // Full article content (mock)
  const articleContent = {
    introduction: newsItem.excerpt || "Aventra Consulting Inc., a global IT and business consulting services leader, today announced significant developments that underscore the company's commitment to innovation and client success.",
    
    mainContent: [
      "The announcement comes as part of Aventra Consulting's ongoing strategy to enhance its service delivery capabilities and strengthen its position as a trusted technology partner for organizations worldwide. This development reflects the company's continued investment in cutting-edge technology and expertise.",
      
      "\"At Aventra Consulting, we are committed to delivering exceptional value to our clients through innovation, expertise, and a deep understanding of their business challenges,\" said Vikas Grover, President and Chief Executive Officer of Aventra Consulting. \"This announcement represents another significant milestone in our journey to help clients navigate the complexities of digital transformation.\"",
      
      "The initiative builds upon Aventra Consulting's proven track record of successfully implementing large-scale technology solutions across multiple industries. With over 100+ professionals operating in more than 20 locations worldwide, Aventra Consulting brings unparalleled scale and expertise to every client engagement.",
      
      "Industry analysts have responded positively to the announcement, noting Aventra Consulting's consistent ability to anticipate market trends and position itself strategically for future growth. The company's member-owner culture and proximity model continue to differentiate it in an increasingly competitive marketplace.",
      
      "Aventra Consulting's clients span virtually every industry sector, including government, financial services, healthcare, manufacturing, retail, and energy and utilities. The company's comprehensive service portfolio encompasses business consulting, systems integration, managed IT services, and intellectual property-based solutions."
    ],
    
    quote: {
      text: "This development reinforces Aventra Consulting's position as a leader in the global IT services market and demonstrates our unwavering commitment to innovation and client success.",
      author: "Rovil Mahajan",
      title: "Executive Vice-President and CFO, Aventra Consulting"
    },
    
    aboutCGI: "Founded in 2011, Aventra Consulting is among the largest independent IT and business consulting services firms in the world. Operating in hundreds of locations across the globe, Aventra Consulting delivers an end-to-end portfolio of capabilities, from strategic IT and business consulting to systems integration, managed IT services, and intellectual property solutions. Aventra Consulting works with clients through a local relationship model complemented by a global delivery network that helps clients digitally transform their organizations and accelerate results.",
    
    forwardLooking: "This news release contains forward-looking statements about future events and circumstances. Actual results may differ materially from those expressed or implied in such statements. Factors that could cause differences include economic conditions, market dynamics, technological changes, and other risks detailed in Aventra Consulting's filings with securities regulators."
  };

  return (
    <div className="min-h-screen">
      {/* Back Navigation */}
      <section className="bg-white border-b border-gray-200 sticky top-16 z-40">
        <div className="max-w-4xl mx-auto px-6 py-4">
          <Link
            to="/news"
            className="inline-flex items-center text-gray-700 hover:text-[#E31937] transition-colors"
          >
            <ArrowLeft className="h-4 w-4 mr-2" />
            Back to News
          </Link>
        </div>
      </section>

      {/* Article Header */}
      <article className="py-12 bg-white">
        <div className="max-w-4xl mx-auto px-6">
          {/* Type Badge */}
          <div className="mb-6">
            <span className="inline-flex items-center gap-2 px-4 py-2 bg-[#E31937] text-white text-sm font-semibold rounded uppercase">
              <FileText className="h-4 w-4" />
              {newsItem.type}
            </span>
          </div>

          {/* Title */}
          <h1 className="text-4xl md:text-5xl font-bold text-gray-900 mb-6 leading-tight">
            {newsItem.title}
          </h1>

          {/* Meta Information */}
          <div className="flex flex-wrap items-center gap-6 text-gray-600 mb-8 pb-8 border-b border-gray-200">
            <div className="flex items-center gap-2">
              <Calendar className="h-5 w-5 text-[#E31937]" />
              <span className="font-medium">{newsItem.date}</span>
            </div>
            <div>
              <span className="text-sm text-gray-600">Gurugram , India</span>
            </div>
          </div>

          {/* Featured Image */}
          <div className="aspect-video mb-12 rounded-lg overflow-hidden shadow-lg">
            <img
              src={newsItem.image}
              alt={newsItem.title}
              className="w-full h-full object-cover"
            />
          </div>

          {/* Article Content */}
          <div className="prose prose-lg max-w-none">
            {/* Introduction */}
            <p className="text-xl text-gray-700 leading-relaxed mb-8 font-medium">
              {articleContent.introduction}
            </p>

            {/* Main Content */}
            {articleContent.mainContent.map((paragraph, index) => (
              <p key={index} className="text-lg text-gray-700 leading-relaxed mb-6">
                {paragraph}
              </p>
            ))}

            {/* Quote Box */}
            <div className="my-12 p-8 bg-[#F8F8F8] border-l-4 border-[#E31937] rounded-lg">
              <blockquote className="text-2xl font-medium text-gray-900 mb-4 italic">
                "{articleContent.quote.text}"
              </blockquote>
              <div className="text-gray-700">
                <div className="font-semibold">{articleContent.quote.author}</div>
                <div className="text-sm">{articleContent.quote.title}</div>
              </div>
            </div>

            {/* About Aventra Consulting */}
            <div className="mt-12 p-6 bg-gray-50 rounded-lg">
              <h3 className="text-xl font-bold text-gray-900 mb-3">About Aventra Consulting</h3>
              <p className="text-gray-700 leading-relaxed">
                {articleContent.aboutCGI}
              </p>
            </div>

            {/* Forward Looking Statement */}
            <div className="mt-8 p-4 bg-yellow-50 border border-yellow-200 rounded">
              <h4 className="text-sm font-bold text-gray-900 mb-2">Forward-Looking Information</h4>
              <p className="text-sm text-gray-700 leading-relaxed">
                {articleContent.forwardLooking}
              </p>
            </div>
          </div>

          {/* Share Section */}
          <div className="mt-12 pt-8 border-t border-gray-200">
            <div className="flex items-center gap-4">
              <span className="text-gray-700 font-medium flex items-center gap-2">
                <Share2 className="h-5 w-5" />
                Share this news:
              </span>
              <div className="flex gap-3">
                <a
                  href="#"
                  className="p-2 text-gray-600 hover:text-[#0077B5] hover:bg-gray-100 rounded transition-colors"
                  aria-label="Share on LinkedIn"
                >
                  <Linkedin className="h-5 w-5" />
                </a>
                <a
                  href="#"
                  className="p-2 text-gray-600 hover:text-[#1DA1F2] hover:bg-gray-100 rounded transition-colors"
                  aria-label="Share on Twitter"
                >
                  <Twitter className="h-5 w-5" />
                </a>
                <a
                  href="#"
                  className="p-2 text-gray-600 hover:text-[#1877F2] hover:bg-gray-100 rounded transition-colors"
                  aria-label="Share on Facebook"
                >
                  <Facebook className="h-5 w-5" />
                </a>
              </div>
            </div>
          </div>

          {/* Media Contact */}
          <div className="mt-12 p-6 bg-[#F8F8F8] rounded-lg">
            <h3 className="text-xl font-bold text-gray-900 mb-4">Media inquiries</h3>
            <div className="space-y-2 text-gray-700">
              <p>
                <strong>Email:</strong>{' '}
                <a href="mailto:media@cgi.com" className="text-[#E31937] hover:underline">
                  sales@consultaventra.com
                </a>
              </p>
              <p>
                <strong>Phone:</strong> +91 9070030003
              </p>
            </div>
          </div>
        </div>
      </article>

      {/* Related News */}
      {relatedNews.length > 0 && (
        <section className="py-16 bg-[#F8F8F8]">
          <div className="max-w-7xl mx-auto px-6">
            <h2 className="text-3xl font-bold text-gray-900 mb-8">Related News</h2>
            <div className="grid md:grid-cols-3 gap-8">
              {relatedNews.map((related) => (
                <Link
                  key={related.id}
                  to={related.link}
                  className="group bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-shadow"
                >
                  <div className="aspect-video overflow-hidden">
                    <img
                      src={related.image}
                      alt={related.title}
                      className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                    />
                  </div>
                  <div className="p-6">
                    <div className="flex items-center justify-between mb-3">
                      <span className="inline-flex items-center gap-1 px-2 py-1 bg-gray-100 text-gray-700 text-xs font-medium rounded uppercase">
                        <FileText className="h-3 w-3" />
                        {related.type}
                      </span>
                      <span className="text-sm text-gray-600">{related.date}</span>
                    </div>
                    <h3 className="text-lg font-bold text-gray-900 group-hover:text-[#E31937] transition-colors line-clamp-2">
                      {related.title}
                    </h3>
                  </div>
                </Link>
              ))}
            </div>
          </div>
        </section>
      )}

      {/* CTA Section */}
      <section className="py-16 bg-[#1A1A1A] text-white">
        <div className="max-w-4xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold mb-4">Get the latest Aventra Consulting news</h2>
          <p className="text-xl text-gray-300 mb-8">
            Subscribe to receive news and updates directly to your inbox.
          </p>
          <form className="flex flex-col sm:flex-row gap-4 max-w-xl mx-auto">
            <input
              type="email"
              placeholder="Enter your email"
              className="flex-1 px-4 py-3 rounded-md text-gray-900 focus:outline-none focus:ring-2 focus:ring-[#E31937]"
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

export default NewsDetailPage;
