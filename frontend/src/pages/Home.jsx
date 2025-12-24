import React from 'react';
import { Link } from 'react-router-dom';
import HeroCarousel from '../components/HeroCarousel';
import { Lightbulb, Network, PuzzlePiece, ArrowRight } from 'lucide-react';
import { heroSlides, industries, services, solutions, caseStudies, experts, news } from '../mockData';

const Home = () => {
  return (
    <div className="min-h-screen">
      {/* Hero Carousel */}
      <HeroCarousel slides={heroSlides} />

      {/* Insights, Services, Solutions Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-3 gap-12">
            {/* Insights */}
            <div className="space-y-4">
              <div className="flex items-center space-x-3 mb-4">
                <div className="p-3 border-2 border-gray-900 rounded-lg">
                  <Lightbulb className="h-8 w-8 text-gray-900" />
                </div>
                <h2 className="text-2xl font-bold text-gray-900">Insights</h2>
              </div>
              <p className="text-gray-700 leading-relaxed">
                We provide a global antenna based on facts, not hype, to help improve returns on your IT and business investments.
              </p>
              <div className="space-y-2 pt-4">
                {industries.slice(0, 6).map((industry) => (
                  <Link
                    key={industry.name}
                    to={industry.link}
                    className="block text-gray-700 hover:text-[#E31937] transition-colors"
                  >
                    {industry.name}
                  </Link>
                ))}
                <Link to="/industries" className="block text-[#E31937] font-semibold hover:underline pt-2">
                  Explore all industries →
                </Link>
              </div>
            </div>

            {/* Services */}
            <div className="space-y-4">
              <div className="flex items-center space-x-3 mb-4">
                <div className="p-3 border-2 border-gray-900 rounded-lg">
                  <Network className="h-8 w-8 text-gray-900" />
                </div>
                <h2 className="text-2xl font-bold text-gray-900">Services</h2>
              </div>
              <p className="text-gray-700 leading-relaxed">
                We help you meet the ever-evolving digital expectations of your customers and citizens through our secure and flexible end-to-end services.
              </p>
              <div className="space-y-2 pt-4">
                {services.slice(0, 6).map((service) => (
                  <Link
                    key={service.name}
                    to={service.link}
                    className="block text-gray-700 hover:text-[#E31937] transition-colors"
                  >
                    {service.name}
                  </Link>
                ))}
                <Link to="/services" className="block text-[#E31937] font-semibold hover:underline pt-2">
                  Explore end-to-end services →
                </Link>
              </div>
            </div>

            {/* Solutions */}
            <div className="space-y-4">
              <div className="flex items-center space-x-3 mb-4">
                <div className="p-3 border-2 border-gray-900 rounded-lg">
                  <PuzzlePiece className="h-8 w-8 text-gray-900" />
                </div>
                <h2 className="text-2xl font-bold text-gray-900">Solutions</h2>
              </div>
              <p className="text-gray-700 leading-relaxed">
                Our intellectual property-based solutions help you accelerate innovation and drive operational excellence.
              </p>
              <div className="space-y-2 pt-4">
                {solutions.slice(0, 6).map((solution) => (
                  <Link
                    key={solution.name}
                    to={solution.link}
                    className="block text-gray-700 hover:text-[#E31937] transition-colors"
                  >
                    {solution.name}
                  </Link>
                ))}
                <Link to="/solutions" className="block text-[#E31937] font-semibold hover:underline pt-2">
                  Explore all solutions →
                </Link>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Case Studies Section */}
      <section className="py-16 bg-gray-50">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-3xl font-bold text-gray-900 mb-8">Case studies</h2>
          <div className="grid md:grid-cols-2 lg:grid-cols-4 gap-6">
            {caseStudies.map((study) => (
              <Link
                key={study.id}
                to={study.link}
                className="group bg-white overflow-hidden shadow-sm hover:shadow-lg transition-shadow"
              >
                <div className="aspect-video overflow-hidden">
                  <img
                    src={study.image}
                    alt={study.title}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <div className="p-4">
                  <p className="text-sm text-gray-600 mb-2">{study.company}</p>
                  <h3 className="text-base font-semibold text-gray-900 group-hover:text-[#E31937] transition-colors">
                    {study.title}
                  </h3>
                </div>
              </Link>
            ))}
          </div>
          <div className="mt-8 text-center">
            <Link
              to="/case-studies"
              className="inline-flex items-center text-[#E31937] font-semibold hover:underline"
            >
              Discover more of our clients' successes
              <ArrowRight className="ml-2 h-5 w-5" />
            </Link>
          </div>
        </div>
      </section>

      {/* Expert Insights Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <p className="text-lg text-gray-700 mb-8 max-w-4xl">
            Through our global insights and local experts, we help build the strategies and deliver the end-to-end services that help clients achieve meaningful and sustainable outcomes.
          </p>
          <div className="grid md:grid-cols-3 gap-8">
            {experts.map((expert) => (
              <Link
                key={expert.id}
                to={expert.link}
                className="group"
              >
                <div className="aspect-square overflow-hidden rounded-full mb-4">
                  <img
                    src={expert.image}
                    alt={expert.name}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <h3 className="font-semibold text-gray-900 mb-2">{expert.name}</h3>
                <p className="text-gray-700 group-hover:text-[#E31937] transition-colors">
                  {expert.title}
                </p>
              </Link>
            ))}
          </div>
          <div className="mt-8 text-center">
            <Link
              to="/blog"
              className="inline-flex items-center text-[#E31937] font-semibold hover:underline"
            >
              Explore advice and commentary from our experts
              <ArrowRight className="ml-2 h-5 w-5" />
            </Link>
          </div>
        </div>
      </section>

      {/* Video Section */}
      <section className="py-16 bg-gray-900">
        <div className="max-w-7xl mx-auto px-6">
          <div className="text-center mb-8">
            <h2 className="text-3xl font-bold text-white mb-4">CGI – by your side, every day</h2>
          </div>
          <div className="max-w-4xl mx-auto aspect-video bg-gray-800 rounded-lg overflow-hidden">
            <iframe
              className="w-full h-full"
              src="https://www.youtube.com/embed/bb62GbVZa8k"
              title="CGI Video"
              frameBorder="0"
              allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
              allowFullScreen
            ></iframe>
          </div>
        </div>
      </section>

      {/* Latest News Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-3xl font-bold text-gray-900 mb-8">Latest news</h2>
          <div className="grid md:grid-cols-3 gap-6">
            {news.map((item) => (
              <Link
                key={item.id}
                to={item.link}
                className="group bg-white overflow-hidden shadow-sm hover:shadow-lg transition-shadow"
              >
                <div className="aspect-video overflow-hidden">
                  <img
                    src={item.image}
                    alt={item.title}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>
                <div className="p-4">
                  <p className="text-sm text-gray-600 mb-2">{item.type}</p>
                  <h3 className="text-base font-semibold text-gray-900 group-hover:text-[#E31937] transition-colors">
                    {item.title}
                  </h3>
                </div>
              </Link>
            ))}
          </div>
          <div className="mt-8 text-center">
            <Link
              to="/news"
              className="inline-flex items-center text-[#E31937] font-semibold hover:underline"
            >
              See more
              <ArrowRight className="ml-2 h-5 w-5" />
            </Link>
          </div>
        </div>
      </section>
    </div>
  );
};

export default Home;
