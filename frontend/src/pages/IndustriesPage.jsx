import React from 'react';
import { Link } from 'react-router-dom';
import { ArrowRight } from 'lucide-react';

const IndustriesPage = () => {
  const industriesData = [
    {
      name: 'Banking',
      description: 'Helping financial institutions modernize operations, enhance customer experiences, and navigate regulatory challenges.',
      image: 'https://images.unsplash.com/photo-1501167786227-4cba60f6d58f?w=600&h=400&fit=crop',
      link: '/industries/banking'
    },
    {
      name: 'Government',
      description: 'Supporting government agencies in delivering better services to citizens through digital transformation and modernization.',
      image: 'https://images.unsplash.com/photo-1479142506502-19b3a3b7ff33?w=600&h=400&fit=crop',
      link: '/industries/government'
    },
    {
      name: 'Health',
      description: 'Transforming healthcare delivery with innovative technology solutions that improve patient outcomes and operational efficiency.',
      image: 'https://images.unsplash.com/photo-1576091160399-112ba8d25d1d?w=600&h=400&fit=crop',
      link: '/industries/health'
    },
    {
      name: 'Manufacturing',
      description: 'Enabling manufacturers to optimize operations, improve supply chain visibility, and embrace Industry 4.0.',
      image: 'https://images.unsplash.com/photo-1565043589221-1a6fd9ae45c7?w=600&h=400&fit=crop',
      link: '/industries/manufacturing'
    },
    {
      name: 'Retail',
      description: 'Helping retailers create seamless omnichannel experiences and adapt to changing consumer expectations.',
      image: 'https://images.unsplash.com/photo-1441986300917-64674bd600d8?w=600&h=400&fit=crop',
      link: '/industries/retail'
    },
    {
      name: 'Energy & utilities',
      description: 'Supporting energy and utility companies in their transition to sustainable, digital-first operations.',
      image: 'https://images.unsplash.com/photo-1473341304170-971dccb5ac1e?w=600&h=400&fit=crop',
      link: '/industries/energy'
    }
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-gradient-to-br from-gray-50 to-gray-100 py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold text-gray-900 mb-6">Industries we serve</h1>
          <p className="text-xl text-gray-700 max-w-3xl">
            We provide industry-specific insights and solutions that help you navigate unique challenges and seize opportunities in your sector.
          </p>
        </div>
      </section>

      {/* Industries Grid */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
            {industriesData.map((industry) => (
              <Link
                key={industry.name}
                to={industry.link}
                className="group bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-shadow"
              >
                <div className="aspect-video overflow-hidden">
                  <img
                    src={industry.image}
                    alt={industry.name}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <div className="p-6">
                  <h3 className="text-2xl font-bold text-gray-900 mb-3 group-hover:text-[#E31937] transition-colors">
                    {industry.name}
                  </h3>
                  <p className="text-gray-700 mb-4">{industry.description}</p>
                  <span className="inline-flex items-center text-[#E31937] font-semibold group-hover:underline">
                    Learn more
                    <ArrowRight className="ml-2 h-5 w-5 group-hover:translate-x-1 transition-transform" />
                  </span>
                </div>
              </Link>
            ))}
          </div>
        </div>
      </section>

      {/* Voice of Clients Section */}
      <section className="py-16 bg-gray-50">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold text-gray-900 mb-4">Voice of our clients</h2>
          <p className="text-xl text-gray-700 mb-8 max-w-3xl mx-auto">
            Every year, we conduct thousands of in-depth client interviews to understand their priorities, challenges, and expectations.
          </p>
          <Link
            to="/voice-of-clients"
            className="inline-flex items-center text-[#E31937] font-semibold hover:underline text-lg"
          >
            Read the latest insights
            <ArrowRight className="ml-2 h-5 w-5" />
          </Link>
        </div>
      </section>
    </div>
  );
};

export default IndustriesPage;
