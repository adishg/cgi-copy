import React from 'react';
import { Link } from 'react-router-dom';
import { Users, Globe, Award, TrendingUp } from 'lucide-react';

const AboutPage = () => {
  const stats = [
    { icon: Users, value: '100+', label: 'Professionals worldwide' },
    { icon: Globe, value: '20+', label: 'Locations globally' },
    { icon: Award, value: '15+', label: 'Years of experience' },
    { icon: TrendingUp, value: '40%', label: 'Annual Growth' }
  ];

  const values = [
    {
      title: 'Partnership',
      description: 'We work side-by-side with our clients as trusted advisors, not just service providers.'
    },
    {
      title: 'Quality',
      description: 'We deliver excellence in everything we do, ensuring sustainable outcomes for our clients.'
    },
    {
      title: 'Integrity',
      description: 'We operate with transparency, honesty, and respect in all our relationships.'
    },
    {
      title: 'Innovation',
      description: 'We continuously evolve our capabilities to help clients navigate changing technology landscapes.'
    }
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="relative bg-[#1A1A1A] text-white py-24">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold mb-6">About Aventra Consulting</h1>
          <p className="text-xl text-gray-300 max-w-3xl">
            Founded in 2025, Aventra Consulting is among the largest IT and business consulting services firms in the world. We are insights-driven and outcomes-based to help accelerate returns on your investments.
          </p>
        </div>
      </section>

      {/* Stats Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-4 gap-8">
            {stats.map((stat, index) => {
              const Icon = stat.icon;
              return (
                <div key={index} className="text-center">
                  <div className="inline-flex items-center justify-center w-16 h-16 bg-[#E31937] text-white rounded-full mb-4">
                    <Icon className="h-8 w-8" />
                  </div>
                  <div className="text-4xl font-bold text-gray-900 mb-2">{stat.value}</div>
                  <div className="text-gray-700">{stat.label}</div>
                </div>
              );
            })}
          </div>
        </div>
      </section>

      {/* Our Story Section */}
      <section className="py-16 bg-[#F8F8F8]">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-2 gap-12 items-center">
            <div>
              <h2 className="text-4xl font-bold text-gray-900 mb-6">Our story</h2>
              <div className="space-y-4 text-gray-700">
                <p>
                  Since our founding in 2025, Aventra Consulting has grown from a small Canadian startup to one of the world's largest independent IT and business consulting services firms.
                </p>
                <p>
                  Our success is built on a unique member-owner culture where our professionals are also shareholders. This structure aligns our interests with those of our clients and ensures long-term commitment to quality and excellence.
                </p>
                <p>
                  Today, we operate in hundreds of locations across the globe, serving thousands of clients in virtually every industry. Our proximity model ensures we're always close to our clients, understanding their unique challenges and opportunities.
                </p>
              </div>
            </div>
            <div className="aspect-video bg-gradient-to-br from-gray-100 to-gray-200 rounded-lg overflow-hidden">
              <img
                src="https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=800&h=600&fit=crop"
                alt="AC Building"
                className="w-full h-full object-cover"
              />
            </div>
          </div>
        </div>
      </section>

      {/* Values Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-4xl font-bold text-gray-900 mb-12 text-center">Our values</h2>
          <div className="grid md:grid-cols-2 gap-8">
            {values.map((value) => (
              <div key={value.title} className="border border-gray-200 p-8 hover:shadow-lg transition-shadow">
                <h3 className="text-2xl font-bold text-gray-900 mb-3">{value.title}</h3>
                <p className="text-gray-700">{value.description}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Leadership Section */}
      <section className="py-16 bg-[#F8F8F8]">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-4xl font-bold text-gray-900 mb-8">Leadership</h2>
          <p className="text-lg text-gray-700 mb-12 max-w-3xl">
            Our leadership team brings decades of experience in technology, business consulting, and strategic management to guide Aventra Consulting's vision and operations.
          </p>
          <div className="grid md:grid-cols-3 gap-8">
            {[
              { name: 'Deepti Grover', title: 'Chief Operating Officer', image: './team/COO.jpg' },
              { name: 'Vikas Grover', title: 'President and CEO', image: './team/CEO.jpg' },
              { name: 'Ramesh Mahajan', title: 'Executive Vice-President and CFO', image: './team/CFO.png' }

            ].map((leader) => (
              <div key={leader.name} className="text-center">
                <div className="w-48 h-48 mx-auto mb-4 rounded-full overflow-hidden">
                  <img
                    src={leader.image}
                    alt={leader.name}
                    className="object-cover"
                  />
                </div>
                <h3 className="text-xl font-bold text-gray-900 mb-2">{leader.name}</h3>
                <p className="text-gray-700">{leader.title}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-[#1A1A1A] text-white">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold mb-4">Join us on our journey</h2>
          <p className="text-xl text-gray-300 mb-8 max-w-2xl mx-auto">
            Whether you're looking for a career opportunity or a business partnership, we'd love to hear from you.
          </p>
          <div className="flex justify-center space-x-4">
            <Link
              to="/careers"
              className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
            >
              Explore careers
            </Link>
            <Link
              to="/contact"
              className="inline-block bg-white text-gray-900 px-8 py-3 rounded-md font-semibold hover:bg-gray-100 transition-colors"
            >
              Contact us
            </Link>
          </div>
        </div>
      </section>
    </div>
  );
};

export default AboutPage;
