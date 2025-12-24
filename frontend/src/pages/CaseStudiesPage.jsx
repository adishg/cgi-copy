import React from 'react';
import { Link } from 'react-router-dom';
import { caseStudies } from '../mockData';
import { ArrowRight } from 'lucide-react';

const CaseStudiesPage = () => {
  const allCaseStudies = [
    ...caseStudies,
    {
      id: 5,
      title: 'Transforming customer experience for a leading telecom provider',
      company: 'Telecommunications',
      image: 'https://images.unsplash.com/photo-1551434678-e076c223a692?w=600&h=400&fit=crop',
      link: '/case-studies/telecom'
    },
    {
      id: 6,
      title: 'Modernizing legacy systems for a global bank',
      company: 'Banking',
      image: 'https://images.unsplash.com/photo-1556742111-a301076d9d18?w=600&h=400&fit=crop',
      link: '/case-studies/banking-legacy'
    },
    {
      id: 7,
      title: 'Building a smart city platform for urban innovation',
      company: 'Government',
      image: 'https://images.unsplash.com/photo-1480714378408-67cf0d13bc1b?w=600&h=400&fit=crop',
      link: '/case-studies/smart-city'
    },
    {
      id: 8,
      title: 'Implementing AI-powered supply chain optimization',
      company: 'Retail',
      image: 'https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=600&h=400&fit=crop',
      link: '/case-studies/supply-chain'
    },
    {
      id: 9,
      title: 'Digital transformation in healthcare delivery',
      company: 'Health',
      image: 'https://images.unsplash.com/photo-1538108149393-fbbd81895907?w=600&h=400&fit=crop',
      link: '/case-studies/healthcare-digital'
    }
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-gradient-to-br from-gray-50 to-gray-100 py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold text-gray-900 mb-6">Client success stories</h1>
          <p className="text-xl text-gray-700 max-w-3xl">
            Discover how we've helped clients across industries achieve meaningful and sustainable outcomes through our expertise and partnership.
          </p>
        </div>
      </section>

      {/* Case Studies Grid */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
            {allCaseStudies.map((study) => (
              <Link
                key={study.id}
                to={study.link}
                className="group bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-shadow"
              >
                <div className="aspect-video overflow-hidden">
                  <img
                    src={study.image}
                    alt={study.title}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <div className="p-6">
                  <p className="text-sm text-[#E31937] font-semibold mb-2">{study.company}</p>
                  <h3 className="text-lg font-bold text-gray-900 mb-3 group-hover:text-[#E31937] transition-colors line-clamp-2">
                    {study.title}
                  </h3>
                  <span className="inline-flex items-center text-[#E31937] font-semibold group-hover:underline">
                    Read more
                    <ArrowRight className="ml-2 h-5 w-5 group-hover:translate-x-1 transition-transform" />
                  </span>
                </div>
              </Link>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-gray-900 text-white">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold mb-4">Ready to write your success story?</h2>
          <p className="text-xl text-gray-300 mb-8 max-w-2xl mx-auto">
            Let's explore how we can help you achieve your business objectives.
          </p>
          <Link
            to="/contact"
            className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
          >
            Get in touch
          </Link>
        </div>
      </section>
    </div>
  );
};

export default CaseStudiesPage;
