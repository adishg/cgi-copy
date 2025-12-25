import React from 'react';
import { Link } from 'react-router-dom';
import { services } from '../mockData';
import { ArrowRight, Check } from 'lucide-react';

const ServicesPage = () => {
  const serviceDetails = [
    {
      name: 'Business consulting',
      description: 'Strategic advisory services that help you navigate complex business challenges and drive meaningful transformation.',
      benefits: ['Strategic planning', 'Digital transformation', 'Process optimization', 'Change management']
    },
    {
      name: 'Business process services',
      description: 'Comprehensive BPS solutions that streamline operations and enhance efficiency across your organization.',
      benefits: ['Process automation', 'Workflow optimization', 'Cost reduction', 'Quality improvement']
    },
    {
      name: 'Managed IT services',
      description: 'End-to-end IT management that allows you to focus on your core business while we handle your technology needs.',
      benefits: ['24/7 support', 'Infrastructure management', 'Security monitoring', 'Performance optimization']
    },
    {
      name: 'Artificial intelligence',
      description: 'AI solutions that transform data into actionable insights and drive intelligent automation across your enterprise.',
      benefits: ['Machine learning', 'Natural language processing', 'Computer vision', 'Predictive analytics']
    },
    {
      name: 'Data analytics',
      description: 'Advanced analytics capabilities that unlock the value of your data and drive data-driven decision making.',
      benefits: ['Business intelligence', 'Advanced analytics', 'Data visualization', 'Real-time insights']
    },
    {
      name: 'Cloud & hybrid IT',
      description: 'Cloud solutions that provide the flexibility, scalability, and security your business needs to thrive.',
      benefits: ['Cloud migration', 'Hybrid cloud', 'Cloud security', 'Cloud optimization']
    }
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-[#F8F8F8] py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold text-gray-900 mb-6">End-to-end services</h1>
          <p className="text-xl text-gray-700 max-w-3xl">
            We help you meet the ever-evolving digital expectations of your customers and citizens through our secure and flexible end-to-end services.
          </p>
        </div>
      </section>

      {/* Services Grid */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="grid gap-12">
            {serviceDetails.map((service, index) => (
              <div key={service.name} className="grid md:grid-cols-2 gap-8 items-center">
                <div className={index % 2 === 1 ? 'md:order-2' : ''}>
                  <h2 className="text-3xl font-bold text-gray-900 mb-4">{service.name}</h2>
                  <p className="text-lg text-gray-700 mb-6">{service.description}</p>
                  <ul className="space-y-2 mb-6">
                    {service.benefits.map((benefit) => (
                      <li key={benefit} className="flex items-center text-gray-700">
                        <Check className="h-5 w-5 text-[#E31937] mr-2 flex-shrink-0" />
                        {benefit}
                      </li>
                    ))}
                  </ul>
                  <Link
                    to={`/services/${service.name.toLowerCase().replace(/ /g, '-')}`}
                    className="inline-flex items-center text-[#E31937] font-semibold hover:underline"
                  >
                    Learn more
                    <ArrowRight className="ml-2 h-5 w-5" />
                  </Link>
                </div>
                <div className={index % 2 === 1 ? 'md:order-1' : ''}>
                  <div className="aspect-video bg-gradient-to-br from-gray-100 to-gray-200 rounded-lg overflow-hidden">
                    <img
                      src={`https://images.unsplash.com/photo-${['1551434678-e076c223a692', '1460925895917-afdab827c52f', '1451187580459-43490279c0fa', '1677442136019-21780ecad995', '1551288049-bebda4e38f71', '1484480974693-6ca0a78fb36b'][index]}?w=800&h=600&fit=crop`}
                      alt={service.name}
                      className="w-full h-full object-cover"
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-[#1A1A1A] text-white">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold mb-4">Ready to transform your business?</h2>
          <p className="text-xl text-gray-300 mb-8 max-w-2xl mx-auto">
            Let's discuss how our services can help you achieve your business objectives.
          </p>
          <Link
            to="/contact"
            className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
          >
            Contact us
          </Link>
        </div>
      </section>
    </div>
  );
};

export default ServicesPage;
