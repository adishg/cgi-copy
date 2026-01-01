import React from 'react';
import { Link, useParams } from 'react-router-dom';
import { blogPosts } from '../mockData';
import { Calendar, Clock, User, ArrowLeft, Tag, Share2, Linkedin, Twitter, Facebook } from 'lucide-react';

const BlogDetailPage = () => {
  const { blogId } = useParams();
  
  // Find the blog post by ID or use first post as default
  const post = blogPosts.find(p => p.link.includes(blogId)) || blogPosts[0];
  
  // Get related posts (exclude current post)
  const relatedPosts = blogPosts.filter(p => p.id !== post.id).slice(0, 3);

  // Full article content (mock)
  const articleContent = {
    introduction: "In today's rapidly evolving technological landscape, artificial intelligence has transcended from a buzzword to a fundamental business imperative. Organizations worldwide are investing billions in AI initiatives, yet many struggle to translate these investments into tangible business outcomes. This comprehensive analysis explores the practical implementation strategies that separate successful AI adopters from those still searching for value.",
    
    sections: [
      {
        title: "Understanding the AI Maturity Curve",
        content: "Organizations typically progress through distinct phases in their AI journey: experimentation, pilot programs, scaled deployment, and optimization. Understanding where your organization sits on this curve is crucial for setting realistic expectations and allocating resources effectively. Many enterprises make the mistake of attempting to scale before establishing solid foundations in data quality and organizational readiness."
      },
      {
        title: "The Data Foundation Challenge",
        content: "High-quality AI outcomes require high-quality data inputs. Organizations must invest in robust data governance frameworks, establish clear data ownership, and implement stringent quality control measures. The most successful AI implementations begin not with algorithms, but with a comprehensive data strategy that addresses collection, storage, processing, and accessibility across the enterprise."
      },
      {
        title: "Ethical Considerations and Governance",
        content: "As AI systems increasingly influence business decisions, establishing ethical guidelines and governance frameworks becomes paramount. Organizations must address bias in training data, ensure transparency in algorithmic decision-making, and maintain human oversight for critical processes. Building trust with stakeholders requires demonstrating responsible AI practices from the outset."
      },
      {
        title: "Building Cross-Functional AI Teams",
        content: "Successful AI initiatives require collaboration between data scientists, domain experts, IT professionals, and business leaders. Organizations should invest in upskilling existing workforce while strategically hiring specialized talent. Creating a culture that embraces experimentation and learns from failures is equally important as technical capabilities."
      },
      {
        title: "Measuring ROI and Business Impact",
        content: "Establishing clear metrics for AI success goes beyond technical performance indicators. Organizations must define business KPIs that demonstrate tangible value: revenue growth, cost reduction, improved customer satisfaction, or operational efficiency gains. Regular assessment and course correction ensure AI investments deliver measurable returns."
      }
    ],
    
    conclusion: "The path from AI experimentation to enterprise-wide value creation requires strategic thinking, robust infrastructure, and organizational commitment. By focusing on practical implementation over technological hype, addressing ethical considerations proactively, and maintaining clear business alignment, organizations can unlock AI's transformative potential while managing associated risks effectively."
  };

  return (
    <div className="min-h-screen">
      {/* Back Navigation */}
      <section className="bg-white border-b border-gray-200 sticky top-16 z-40">
        <div className="max-w-4xl mx-auto px-6 py-4">
          <Link
            to="/blog"
            className="inline-flex items-center text-gray-700 hover:text-[#E31937] transition-colors"
          >
            <ArrowLeft className="h-4 w-4 mr-2" />
            Back to Blog
          </Link>
        </div>
      </section>

      {/* Article Header */}
      <article className="py-12 bg-white">
        <div className="max-w-4xl mx-auto px-6">
          {/* Category Badge */}
          <div className="mb-6">
            <span className="inline-block px-4 py-2 bg-[#E31937] text-white text-sm font-semibold rounded">
              {post.category}
            </span>
          </div>

          {/* Title */}
          <h1 className="text-4xl md:text-5xl font-bold text-gray-900 mb-6 leading-tight">
            {post.title}
          </h1>

          {/* Meta Information */}
          <div className="flex flex-wrap items-center gap-4 md:gap-6 text-gray-600 mb-8 pb-8 border-b border-gray-200">
            {/* <div className="flex items-center gap-2">
              <User className="h-5 w-5 text-[#E31937]" />
              <div>
                <div className="font-semibold text-gray-900">{post.author}</div>
                <div className="text-sm">{post.authorRole}</div>
              </div>
            </div> */}
            <div className="flex items-center gap-2">
              <Calendar className="h-5 w-5 text-[#E31937]" />
              <span>{post.date}</span>
            </div>
            <div className="flex items-center gap-2">
              <Clock className="h-5 w-5 text-[#E31937]" />
              <span>{post.readTime}</span>
            </div>
          </div>

          {/* Featured Image */}
          <div className="aspect-video mb-12 rounded-lg overflow-hidden">
            <img
              src={post.image}
              alt={post.title}
              className="w-full h-full object-cover"
            />
          </div>

          {/* Article Content */}
          <div className="prose prose-lg max-w-none">
            {/* Introduction */}
            <p className="text-xl text-gray-700 leading-relaxed mb-8 font-medium">
              {articleContent.introduction}
            </p>

            {/* Sections */}
            {articleContent.sections.map((section, index) => (
              <div key={index} className="mb-8">
                <h2 className="text-3xl font-bold text-gray-900 mb-4">
                  {section.title}
                </h2>
                <p className="text-lg text-gray-700 leading-relaxed">
                  {section.content}
                </p>
              </div>
            ))}

            {/* Conclusion */}
            <div className="mt-12 p-6 bg-[#F8F8F8] border-l-4 border-[#E31937] rounded">
              <h3 className="text-2xl font-bold text-gray-900 mb-4">Key Takeaway</h3>
              <p className="text-lg text-gray-700 leading-relaxed">
                {articleContent.conclusion}
              </p>
            </div>
          </div>

          {/* Tags */}
          <div className="flex flex-wrap gap-2 mt-12 pt-8 border-t border-gray-200">
            <span className="text-gray-700 font-medium">Tags:</span>
            {post.tags.map((tag) => (
              <span
                key={tag}
                className="inline-flex items-center gap-1 px-3 py-1 bg-gray-100 text-gray-700 rounded-full hover:bg-gray-200 transition-colors cursor-pointer"
              >
                <Tag className="h-3 w-3" />
                {tag}
              </span>
            ))}
          </div>

          {/* Share Section */}
          <div className="mt-8 pt-8 border-t border-gray-200">
            <div className="flex items-center gap-4">
              <span className="text-gray-700 font-medium flex items-center gap-2">
                <Share2 className="h-5 w-5" />
                Share this article:
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

          {/* Author Bio */}
          {/* <div className="mt-12 p-6 bg-[#F8F8F8] rounded-lg">
            <div className="flex items-start gap-4">
              <img
                src={`https://randomuser.me/api/portraits/women/${post.id * 10}.jpg`}
                alt={post.author}
                className="w-20 h-20 rounded-full"
              />
              <div>
                <h3 className="text-xl font-bold text-gray-900 mb-1">{post.author}</h3>
                <p className="text-[#E31937] font-semibold mb-2">{post.authorRole}</p>
                <p className="text-gray-700">
                  {post.author} is a thought leader in {post.category.toLowerCase()} with over 15 years of experience helping organizations navigate digital transformation and achieve measurable business outcomes.
                </p>
              </div>
            </div>
          </div> */}
        </div>
      </article>

      {/* Related Articles */}
      {relatedPosts.length > 0 && (
        <section className="py-16 bg-[#F8F8F8]">
          <div className="max-w-7xl mx-auto px-6">
            <h2 className="text-3xl font-bold text-gray-900 mb-8">Related Articles</h2>
            <div className="grid md:grid-cols-3 gap-8">
              {relatedPosts.map((relatedPost) => (
                <Link
                  key={relatedPost.id}
                  to={relatedPost.link}
                  className="group bg-white border border-gray-200 overflow-hidden hover:shadow-xl transition-shadow"
                >
                  <div className="aspect-video overflow-hidden">
                    <img
                      src={relatedPost.image}
                      alt={relatedPost.title}
                      className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                    />
                  </div>
                  <div className="p-6">
                    <span className="text-sm text-[#E31937] font-semibold">
                      {relatedPost.category}
                    </span>
                    <h3 className="text-xl font-bold text-gray-900 mt-2 mb-3 group-hover:text-[#E31937] transition-colors line-clamp-2">
                      {relatedPost.title}
                    </h3>
                    <p className="text-gray-700 mb-4 line-clamp-2">
                      {relatedPost.excerpt}
                    </p>
                    <div className="flex items-center gap-2 text-sm text-gray-600">
                      <Clock className="h-4 w-4" />
                      {relatedPost.readTime}
                    </div>
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
          <h2 className="text-3xl font-bold mb-4">Ready to transform your business?</h2>
          <p className="text-xl text-gray-300 mb-8">
            Connect with our experts to discuss how we can help you achieve your goals.
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

export default BlogDetailPage;
