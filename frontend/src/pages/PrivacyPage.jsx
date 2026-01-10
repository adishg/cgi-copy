import React from 'react';

const PrivacyPage = () => {
  const lastUpdated = 'January 6, 2026';

  const sections = [
    {
      id: 'introduction',
      title: 'Introduction',
      content: `Aventra Consulting Inc. and its affiliated companies ("Aventra Consulting", "we", "us", or "our") are committed to protecting your privacy. This Privacy Policy explains how we collect, use, disclose, and safeguard your personal information when you visit our websites, use our services, or interact with us in any way.

We respect your privacy and are committed to maintaining the confidentiality and security of personal information entrusted to us. This policy applies to all personal information collected by Aventra Consulting, whether online, offline, or through any other means.`
    },
    {
      id: 'information-we-collect',
      title: 'Information We Collect',
      content: `We may collect and process the following types of personal information:

**Contact Information:** Name, email address, phone number, mailing address, and other contact details you provide when you reach out to us, subscribe to our newsletters, or request information about our services.

**Professional Information:** Job title, company name, industry, and professional background when you engage with our business services or attend our events.

**Website Usage Data:** Information about how you use our website, including your IP address, browser type, operating system, pages visited, time spent on pages, and referring website addresses.

**Communication Records:** Records of your communications with us, including emails, phone calls, and chat messages.

**Application Information:** If you apply for employment with us, we may collect your resume, cover letter, work history, education, and other information relevant to your application.

**Marketing Preferences:** Your preferences for receiving marketing communications and how you want us to contact you.`
    },
    {
      id: 'how-we-use-information',
      title: 'How We Use Your Information',
      content: `We use the personal information we collect for the following purposes:

**Service Delivery:** To provide our consulting, IT, and business services to you or your organization.

**Communication:** To respond to your inquiries, send you information you've requested, and communicate with you about our services, events, and updates.

**Business Operations:** To manage our business operations, including administration, billing, and contract management.

**Marketing:** To send you marketing communications about our services, events, and insights that may be of interest to you, where permitted by law.

**Website Improvement:** To analyze website usage and improve our website's functionality, user experience, and content.

**Legal Compliance:** To comply with applicable laws, regulations, and legal processes.

**Security:** To protect the security and integrity of our systems, networks, and data.

**Recruitment:** To process job applications and manage our recruitment activities.`
    },
    {
      id: 'information-sharing',
      title: 'Information Sharing and Disclosure',
      content: `We may share your personal information in the following circumstances:

**Affiliated Companies:** We may share information with our affiliated companies for internal business purposes, including providing services and support.

**Service Providers:** We may share information with third-party service providers who perform services on our behalf, such as hosting providers, analytics providers, and marketing service providers. These providers are contractually obligated to protect your information.

**Business Partners:** With your consent, we may share information with business partners for joint marketing initiatives or service delivery.

**Legal Requirements:** We may disclose information when required by law, court order, or governmental authority, or when we believe disclosure is necessary to protect our rights, safety, or property.

**Business Transfers:** In the event of a merger, acquisition, or sale of assets, your information may be transferred to the acquiring entity.

We do not sell your personal information to third parties.`
    },
    {
      id: 'data-security',
      title: 'Data Security',
      content: `We implement appropriate technical and organizational security measures to protect your personal information against unauthorized access, alteration, disclosure, or destruction. These measures include:

**Technical Safeguards:** Encryption, firewalls, secure servers, and access controls to protect data during transmission and storage.

**Administrative Safeguards:** Employee training, access restrictions, and security policies to ensure proper handling of personal information.

**Physical Safeguards:** Secure facilities and controlled access to protect physical records and systems.

While we strive to protect your personal information, no method of transmission over the internet or electronic storage is completely secure. We cannot guarantee absolute security of your data.`
    },
    {
      id: 'data-retention',
      title: 'Data Retention',
      content: `We retain your personal information for as long as necessary to fulfill the purposes for which it was collected, including to satisfy legal, accounting, or reporting requirements.

The retention period may vary depending on the context of the processing and our legal obligations. When personal information is no longer required, we securely delete or anonymize it in accordance with our data retention policies.`
    },
    {
      id: 'your-rights',
      title: 'Your Rights and Choices',
      content: `Depending on your location and applicable laws, you may have certain rights regarding your personal information:

**Access:** You may request access to the personal information we hold about you.

**Correction:** You may request that we correct inaccurate or incomplete personal information.

**Deletion:** You may request that we delete your personal information, subject to certain exceptions.

**Objection:** You may object to certain processing of your personal information.

**Portability:** You may request a copy of your personal information in a structured, commonly used format.

**Marketing Opt-Out:** You may opt out of receiving marketing communications from us at any time by clicking the "unsubscribe" link in our emails or contacting us directly.

To exercise any of these rights, please contact us using the contact information provided below.`
    },
    {
      id: 'cookies',
      title: 'Cookies and Tracking Technologies',
      content: `Our website uses cookies and similar tracking technologies to enhance your browsing experience, analyze website traffic, and personalize content.

**Essential Cookies:** Required for the website to function properly and cannot be disabled.

**Analytics Cookies:** Help us understand how visitors interact with our website by collecting information anonymously.

**Functional Cookies:** Enable enhanced functionality and personalization.

**Marketing Cookies:** Used to deliver relevant advertisements and track the effectiveness of marketing campaigns.

You can manage your cookie preferences through your browser settings. Please note that disabling certain cookies may affect the functionality of our website.`
    },
    {
      id: 'international-transfers',
      title: 'International Data Transfers',
      content: `As a global company, Aventra Consulting may transfer your personal information to countries other than the one in which you reside. We ensure that appropriate safeguards are in place to protect your information when it is transferred internationally, including standard contractual clauses and other transfer mechanisms approved by relevant data protection authorities.`
    },
    {
      id: 'children-privacy',
      title: "Children's Privacy",
      content: `Our services are not directed to individuals under the age of 16. We do not knowingly collect personal information from children. If you become aware that a child has provided us with personal information, please contact us, and we will take steps to delete such information.`
    },
    {
      id: 'policy-updates',
      title: 'Changes to This Privacy Policy',
      content: `We may update this Privacy Policy from time to time to reflect changes in our practices or for other operational, legal, or regulatory reasons. We will notify you of any material changes by posting the updated policy on our website with a new "Last Updated" date.

We encourage you to review this Privacy Policy periodically to stay informed about how we are protecting your information.`
    },
    {
      id: 'contact-us',
      title: 'Contact Us',
      content: `If you have any questions, concerns, or requests regarding this Privacy Policy or our data practices, please contact us at:

**Aventra Consulting Inc.**
Privacy Office
Email: privacy@aventraconsulting.com
Phone: +91 98718 33119

We will respond to your inquiry within a reasonable timeframe and in accordance with applicable laws.`
    }
  ];

  return (
    <div className="min-h-screen">
      {/* Hero Section */}
      <section className="bg-[#1A1A1A] text-white py-20">
        <div className="max-w-7xl mx-auto px-6">
          <h1 className="text-5xl font-bold mb-6">Privacy</h1>
          <p className="text-xl text-gray-300 max-w-3xl">
            Aventra Consulting is committed to protecting your privacy and ensuring the security of your personal information.
          </p>
        </div>
      </section>

      {/* Last Updated */}
      <section className="bg-[#F8F8F8] py-6">
        <div className="max-w-7xl mx-auto px-6">
          <p className="text-gray-600">
            <strong>Last Updated:</strong> {lastUpdated}
          </p>
        </div>
      </section>

      {/* Table of Contents */}
      <section className="py-12 bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-6">
          <h2 className="text-2xl font-bold text-gray-900 mb-6">Contents</h2>
          <nav className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
            {sections.map((section, index) => (
              <a
                key={section.id}
                href={`#${section.id}`}
                className="text-gray-700 hover:text-[#E31937] transition-colors"
              >
                {index + 1}. {section.title}
              </a>
            ))}
          </nav>
        </div>
      </section>

      {/* Privacy Policy Content */}
      <section className="py-16 bg-white">
        <div className="max-w-7xl mx-auto px-6">
          <div className="max-w-4xl">
            {sections.map((section, index) => (
              <div
                key={section.id}
                id={section.id}
                className="mb-12 scroll-mt-24"
              >
                <h2 className="text-2xl font-bold text-gray-900 mb-4">
                  {index + 1}. {section.title}
                </h2>
                <div className="prose prose-lg text-gray-700 whitespace-pre-line">
                  {section.content.split('\n\n').map((paragraph, pIndex) => {
                    // Check if paragraph contains bold markers
                    if (paragraph.includes('**')) {
                      const parts = paragraph.split(/(\*\*[^*]+\*\*)/g);
                      return (
                        <p key={pIndex} className="mb-4">
                          {parts.map((part, partIndex) => {
                            if (part.startsWith('**') && part.endsWith('**')) {
                              return (
                                <strong key={partIndex}>
                                  {part.slice(2, -2)}
                                </strong>
                              );
                            }
                            return part;
                          })}
                        </p>
                      );
                    }
                    return (
                      <p key={pIndex} className="mb-4">
                        {paragraph}
                      </p>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-16 bg-[#F8F8F8]">
        <div className="max-w-7xl mx-auto px-6 text-center">
          <h2 className="text-3xl font-bold text-gray-900 mb-4">Questions about privacy?</h2>
          <p className="text-xl text-gray-700 mb-8 max-w-2xl mx-auto">
            If you have any questions or concerns about our privacy practices, we are here to help.
          </p>
          <a
            href="mailto:privacy@aventraconsulting.com"
            className="inline-block bg-[#E31937] text-white px-8 py-3 rounded-md font-semibold hover:bg-[#C01530] transition-colors"
          >
            Contact Privacy Office
          </a>
        </div>
      </section>
    </div>
  );
};

export default PrivacyPage;
