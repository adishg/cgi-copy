import React from 'react';
import { Link } from 'react-router-dom';
import { careerRoles } from '../mockData';
import { MapPin, Clock, ArrowRight } from 'lucide-react';

const CareersPage = () => {
  const benefits = [
    'Competitive compensation and benefits',
    'Flexible work arrangements',
    'Professional development opportunities',
    'Collaborative and inclusive culture',
    'Global career mobility',
    'Work-life balance support'
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="relative bg-gradient-to-br from-gray-900 to-gray-800 text-white py-24">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold mb-6">Build your career with CGI</h1>
          <p className="text-xl text-gray-300 max-w-3xl">
            Join a global team of professionals dedicated to making a difference. At CGI, we're committed to your growth, innovation, and success.
          </p>
        </div>
      </section>

      {/* Why CGI Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-2 gap-12 items-center">
            <div>
              <h2 className="text-4xl font-bold text-gray-900 mb-6">Why CGI?</h2>
              <p className="text-lg text-gray-700 mb-6">
                As one of the world's largest IT and business consulting firms, we offer unparalleled opportunities for professional growth and development. Our member-owner culture ensures that everyone has a voice and a stake in our success.
              </p>
              <ul className="space-y-3">
                {benefits.map((benefit) => (
                  <li key={benefit} className="flex items-start">
                    <div className="w-2 h-2 bg-[#E31937] rounded-full mt-2 mr-3 flex-shrink-0"></div>
                    <span className="text-gray-700">{benefit}</span>
                  </li>
                ))}
              </ul>
            </div>
            <div className="aspect-video bg-gradient-to-br from-gray-100 to-gray-200 rounded-lg overflow-hidden">
              <img
                src="https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=800&h=600&fit=crop"
                alt="CGI Team"
                className="w-full h-full object-cover"
              />
            </div>
          </div>
        </div>
      </section>

      {/* Open Roles Section */}
      <section className="py-16 bg-gray-50">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-4xl font-bold text-gray-900 mb-8">Open positions</h2>
          <div className="grid gap-6">
            {careerRoles.map((role) => (
              <div
                key={role.id}
                className="bg-white border border-gray-200 p-6 hover:shadow-lg transition-shadow"
              >
                <div className="flex flex-col md:flex-row md:items-center md:justify-between">
                  <div className="flex-1 mb-4 md:mb-0">
                    <h3 className="text-2xl font-bold text-gray-900 mb-2">{role.title}</h3>
                    <p className="text-gray-700 mb-4">{role.description}</p>
                    <div className="flex flex-wrap gap-4 text-sm text-gray-600">
                      <span className="flex items-center">
                        <MapPin className="h-4 w-4 mr-1" />
                        {role.location}
                      </span>
                      <span className="flex items-center">
                        <Clock className="h-4 w-4 mr-1" />
                        {role.type}
                      </span>
                    </div>
                  </div>
                  <Link
                    to={`/careers/${role.id}`}
                    className="inline-flex items-center bg-[#E31937] text-white px-6 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors whitespace-nowrap"
                  >
                    Apply now
                    <ArrowRight className="ml-2 h-5 w-5" />
                  </Link>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold text-gray-900 mb-4">Can't find the right role?</h2>
          <p className="text-xl text-gray-700 mb-8 max-w-2xl mx-auto">
            Submit your resume and we'll keep you in mind for future opportunities that match your skills and interests.
          </p>
          <Link
            to="/careers/submit-resume"
            className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
          >
            Submit your resume
          </Link>
        </div>
      </section>
    </div>
  );
};

export default CareersPage;
