import React from 'react';
import { Link } from 'react-router-dom';
import { teamMembers } from '../mockData';
import { Linkedin, Mail } from 'lucide-react';

const TeamPage = () => {
  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-[#1A1A1A] text-white py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold mb-6">Our Leadership Team</h1>
          <p className="text-xl text-gray-300 max-w-3xl">
            Meet the experienced leaders driving CGI's vision, innovation, and commitment to delivering exceptional value to our clients worldwide.
          </p>
        </div>
      </section>

      {/* Team Stats */}
      <section className="py-12 bg-[#F8F8F8]">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-4 gap-8 text-center">
            <div>
              <div className="text-4xl font-bold text-[#E31937] mb-2">94,000+</div>
              <div className="text-gray-700">Team Members Globally</div>
            </div>
            <div>
              <div className="text-4xl font-bold text-[#E31937] mb-2">40+</div>
              <div className="text-gray-700">Countries</div>
            </div>
            <div>
              <div className="text-4xl font-bold text-[#E31937] mb-2">400+</div>
              <div className="text-gray-700">Locations</div>
            </div>
            <div>
              <div className="text-4xl font-bold text-[#E31937] mb-2">48+</div>
              <div className="text-gray-700">Years of Excellence</div>
            </div>
          </div>
        </div>
      </section>

      {/* Leadership Grid */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="mb-12">
            <h2 className="text-3xl font-bold text-gray-900 mb-4">Executive Leadership</h2>
            <p className="text-lg text-gray-700 max-w-3xl">
              Our leadership team brings decades of experience in technology, consulting, and strategic management to guide CGI's global operations and client partnerships.
            </p>
          </div>

          <div className="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-8">
            {teamMembers.map((member) => (
              <div
                key={member.id}
                className="bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-all duration-300 group"
              >
                {/* Photo */}
                <div className="aspect-square overflow-hidden bg-gray-100">
                  <img
                    src={member.image}
                    alt={member.name}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>

                {/* Content */}
                <div className="p-6">
                  <h3 className="text-xl font-bold text-gray-900 mb-2">
                    {member.name}
                  </h3>
                  <p className="text-sm text-[#E31937] font-semibold mb-3">
                    {member.designation}
                  </p>
                  <p className="text-sm text-gray-700 leading-relaxed mb-4">
                    {member.bio}
                  </p>

                  {/* Social Links */}
                  <div className="flex items-center gap-3 pt-4 border-t border-gray-200">
                    <a
                      href="#"
                      className="p-2 text-gray-600 hover:text-[#E31937] hover:bg-gray-100 rounded transition-colors"
                      aria-label={`${member.name} LinkedIn`}
                    >
                      <Linkedin className="h-4 w-4" />
                    </a>
                    <a
                      href="#"
                      className="p-2 text-gray-600 hover:text-[#E31937] hover:bg-gray-100 rounded transition-colors"
                      aria-label={`Email ${member.name}`}
                    >
                      <Mail className="h-4 w-4" />
                    </a>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Member-Owner Culture Section */}
      <section className="py-16 bg-[#F8F8F8]">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid md:grid-cols-2 gap-12 items-center">
            <div>
              <h2 className="text-4xl font-bold text-gray-900 mb-6">
                Our Member-Owner Culture
              </h2>
              <p className="text-lg text-gray-700 mb-4">
                At CGI, we're not just employees—we're member-owners. This unique structure ensures that every professional has a stake in our success and a voice in our future.
              </p>
              <p className="text-lg text-gray-700 mb-6">
                Our leadership team embodies this culture, working alongside our members to drive innovation, deliver excellence, and create lasting value for our clients and stakeholders.
              </p>
              <Link
                to="/about"
                className="inline-block text-[#E31937] font-semibold hover:underline"
              >
                Learn more about CGI →
              </Link>
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

      {/* Diversity & Inclusion */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="text-center max-w-3xl mx-auto mb-12">
            <h2 className="text-4xl font-bold text-gray-900 mb-4">
              Diversity & Inclusion
            </h2>
            <p className="text-lg text-gray-700">
              We believe diverse perspectives drive innovation. Our leadership reflects our commitment to building an inclusive workplace where everyone can thrive and contribute to our shared success.
            </p>
          </div>

          <div className="grid md:grid-cols-3 gap-8">
            <div className="text-center p-6">
              <div className="w-16 h-16 bg-[#E31937] text-white rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-2xl font-bold">40%</span>
              </div>
              <h3 className="text-xl font-bold text-gray-900 mb-2">
                Women in Leadership
              </h3>
              <p className="text-gray-700">
                Committed to gender equity across all levels
              </p>
            </div>

            <div className="text-center p-6">
              <div className="w-16 h-16 bg-[#E31937] text-white rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-2xl font-bold">75+</span>
              </div>
              <h3 className="text-xl font-bold text-gray-900 mb-2">
                Nationalities
              </h3>
              <p className="text-gray-700">
                Global team bringing diverse perspectives
              </p>
            </div>

            <div className="text-center p-6">
              <div className="w-16 h-16 bg-[#E31937] text-white rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-2xl font-bold">100%</span>
              </div>
              <h3 className="text-xl font-bold text-gray-900 mb-2">
                Inclusive Culture
              </h3>
              <p className="text-gray-700">
                Everyone belongs, everyone contributes
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-[#1A1A1A] text-white">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold mb-4">Join Our Team</h2>
          <p className="text-xl text-gray-300 mb-8 max-w-2xl mx-auto">
            Be part of a global team that's shaping the future of technology and business consulting.
          </p>
          <div className="flex flex-col sm:flex-row justify-center gap-4">
            <Link
              to="/careers"
              className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
            >
              Explore Careers
            </Link>
            <Link
              to="/contact"
              className="inline-block bg-white text-gray-900 px-8 py-3 rounded-md font-semibold hover:bg-gray-100 transition-colors"
            >
              Contact Us
            </Link>
          </div>
        </div>
      </section>
    </div>
  );
};

export default TeamPage;
