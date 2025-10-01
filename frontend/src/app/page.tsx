'use client';

import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Sparkles, 
  Download, 
  Eye, 
  Settings, 
  Palette, 
  FileText,
  CheckCircle,
  AlertCircle,
  Loader2,
  ArrowRight,
  Star,
  Zap
} from 'lucide-react';
import { usePresentationStore } from '@/store/presentation';
import ApiService from '@/lib/api';
import { downloadFile } from '@/lib/utils';
import toast from 'react-hot-toast';

const HomePage = () => {
  const {
    currentRequest,
    generationProgress,
    generatedPresentation,
    templates,
    isGenerating,
    error,
    setCurrentRequest,
    setGenerationProgress,
    setGeneratedPresentation,
    setTemplates,
    setIsGenerating,
    setError,
    updateRequestField,
    resetState
  } = usePresentationStore();

  const [step, setStep] = useState(1);
  const [previewSlides, setPreviewSlides] = useState<any[]>([]);
  const [isLoadingPreview, setIsLoadingPreview] = useState(false);

  useEffect(() => {
    // Load templates on component mount
    loadTemplates();
  }, []);

  const loadTemplates = async () => {
    try {
      const response = await ApiService.getTemplates();
      setTemplates(response.data || []);
    } catch (error: any) {
      toast.error('Failed to load templates');
      console.error('Template loading error:', error);
    }
  };

  const handlePreview = async () => {
    if (!currentRequest.title) {
      toast.error('Please enter a presentation title');
      return;
    }

    setIsLoadingPreview(true);
    try {
      const response = await ApiService.previewSlides({
        title: currentRequest.title,
        slides: currentRequest.slides || 5
      });
      
      // Handle the response structure
      const slidesData = response.data?.slides || [];
      setPreviewSlides(slidesData);
      setStep(2);
      toast.success('Preview generated successfully!');
    } catch (error: any) {
      toast.error(`Preview failed: ${error.message}`);
      console.error('Preview error:', error);
    } finally {
      setIsLoadingPreview(false);
    }
  };

  const handleGenerate = async () => {
    if (!currentRequest.title) {
      toast.error('Please enter a presentation title');
      return;
    }

    setIsGenerating(true);
    setError(null);
    
    try {
      // Simulate progress updates
      setGenerationProgress({
        step: 'generating',
        progress: 20,
        message: 'Generating content with AI...'
      });

      const response = await ApiService.generatePresentation(currentRequest as any);
      
      setGenerationProgress({
        step: 'complete',
        progress: 100,
        message: 'Presentation ready!'
      });

      // Handle the response structure properly
      const presentationData = response.data;
      if (presentationData) {
        setGeneratedPresentation(presentationData);
        setStep(3);
        toast.success('Presentation generated successfully!');
      } else {
        throw new Error('No presentation data received');
      }
      
    } catch (error: any) {
      setError(error.message);
      toast.error(`Generation failed: ${error.message}`);
      console.error('Generation error:', error);
    } finally {
      setIsGenerating(false);
      setTimeout(() => setGenerationProgress(null), 2000);
    }
  };

  const handleDownload = async () => {
    if (!generatedPresentation?.presentation_id) {
      toast.error('No presentation to download');
      return;
    }

    try {
      const blob = await ApiService.downloadPresentation(generatedPresentation.presentation_id);
      downloadFile(blob, `${currentRequest.title || 'presentation'}.pptx`);
      toast.success('Download started!');
    } catch (error: any) {
      toast.error(`Download failed: ${error.message}`);
    }
  };

  const handleReset = () => {
    resetState();
    setStep(1);
    setPreviewSlides([]);
    toast.success('Reset to start');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-purple-50 to-pink-50 dark:from-blue-950 dark:via-purple-950 dark:to-pink-950">
      {/* Header */}
      <header className="border-b border-white/20 bg-white/10 backdrop-blur-md">
        <div className="container mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-gradient-to-r from-blue-500 to-purple-600 rounded-lg">
                <Sparkles className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-2xl font-bold gradient-text">AI Presentation Generator</h1>
                <p className="text-sm text-gray-600 dark:text-gray-400">Create stunning presentations with AI</p>
              </div>
            </div>
            
            {step > 1 && (
              <button
                onClick={handleReset}
                className="px-4 py-2 bg-gray-200 dark:bg-gray-700 rounded-lg hover:bg-gray-300 dark:hover:bg-gray-600 transition-colors"
              >
                Start Over
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Progress Indicator */}
      <div className="container mx-auto px-4 py-6">
        <div className="flex items-center justify-center space-x-4 mb-8">
          {[1, 2, 3].map((num) => (
            <div key={num} className="flex items-center">
              <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-semibold transition-colors ${
                step >= num 
                  ? 'bg-blue-500 text-white' 
                  : 'bg-gray-200 dark:bg-gray-700 text-gray-500'
              }`}>
                {step > num ? <CheckCircle className="w-4 h-4" /> : num}
              </div>
              {num < 3 && (
                <div className={`w-16 h-1 mx-2 transition-colors ${
                  step > num ? 'bg-blue-500' : 'bg-gray-200 dark:bg-gray-700'
                }`} />
              )}
            </div>
          ))}
        </div>

        {/* Step Labels */}
        <div className="flex justify-center space-x-20 mb-12 text-sm text-gray-600 dark:text-gray-400">
          <span className={step >= 1 ? 'text-blue-600 dark:text-blue-400 font-semibold' : ''}>
            Create
          </span>
          <span className={step >= 2 ? 'text-blue-600 dark:text-blue-400 font-semibold' : ''}>
            Preview
          </span>
          <span className={step >= 3 ? 'text-blue-600 dark:text-blue-400 font-semibold' : ''}>
            Download
          </span>
        </div>
      </div>

      {/* Main Content */}
      <main className="container mx-auto px-4 pb-12">
        <AnimatePresence mode="wait">
          {step === 1 && (
            <motion.div
              key="step1"
              initial={{ opacity: 0, x: 50 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -50 }}
              className="max-w-4xl mx-auto"
            >
              <div className="glass rounded-2xl p-8 shadow-xl">
                <div className="text-center mb-8">
                  <h2 className="text-3xl font-bold mb-4">Create Your Presentation</h2>
                  <p className="text-gray-600 dark:text-gray-400">
                    Tell us about your presentation and we'll generate it with AI
                  </p>
                </div>

                <div className="grid md:grid-cols-2 gap-8">
                  {/* Left Column - Basic Info */}
                  <div className="space-y-6">
                    <div>
                      <label className="block text-sm font-semibold mb-3 text-gray-700 dark:text-gray-300">
                        Presentation Title *
                      </label>
                      <input
                        type="text"
                        value={currentRequest.title || ''}
                        onChange={(e) => updateRequestField('title', e.target.value)}
                        placeholder="e.g., Introduction to Machine Learning"
                        className="w-full px-4 py-3 rounded-xl border border-gray-200 dark:border-gray-700 bg-white/50 dark:bg-gray-800/50 focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-semibold mb-3 text-gray-700 dark:text-gray-300">
                        Number of Slides
                      </label>
                      <select
                        value={currentRequest.slides || 5}
                        onChange={(e) => updateRequestField('slides', parseInt(e.target.value))}
                        className="w-full px-4 py-3 rounded-xl border border-gray-200 dark:border-gray-700 bg-white/50 dark:bg-gray-800/50 focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                      >
                        {Array.from({ length: 20 }, (_, i) => i + 1).map((num) => (
                          <option key={num} value={num}>
                            {num} slide{num > 1 ? 's' : ''}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-sm font-semibold mb-3 text-gray-700 dark:text-gray-300">
                        Presentation Style
                      </label>
                      <div className="grid grid-cols-2 gap-2">
                        {(['professional', 'creative', 'minimal', 'colorful'] as const).map((style) => (
                          <button
                            key={style}
                            onClick={() => updateRequestField('style', style)}
                            className={`p-3 rounded-lg text-sm font-medium transition-all ${
                              currentRequest.style === style
                                ? 'bg-blue-500 text-white shadow-lg'
                                : 'bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700'
                            }`}
                          >
                            {style.charAt(0).toUpperCase() + style.slice(1)}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div>
                      <label className="block text-sm font-semibold mb-3 text-gray-700 dark:text-gray-300">
                        Tone
                      </label>
                      <div className="grid grid-cols-2 gap-2">
                        {(['formal', 'casual', 'academic', 'business'] as const).map((tone) => (
                          <button
                            key={tone}
                            onClick={() => updateRequestField('tone', tone)}
                            className={`p-3 rounded-lg text-sm font-medium transition-all ${
                              currentRequest.tone === tone
                                ? 'bg-purple-500 text-white shadow-lg'
                                : 'bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700'
                            }`}
                          >
                            {tone.charAt(0).toUpperCase() + tone.slice(1)}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Right Column - Templates */}
                  <div>
                    <label className="block text-sm font-semibold mb-3 text-gray-700 dark:text-gray-300">
                      Choose Template
                    </label>
                    <div className="space-y-3">
                      {templates.map((template) => (
                        <button
                          key={template.id}
                          onClick={() => updateRequestField('template', template.id)}
                          className={`w-full p-4 rounded-xl border-2 text-left transition-all ${
                            currentRequest.template === template.id
                              ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20'
                              : 'border-gray-200 dark:border-gray-700 hover:border-gray-300 dark:hover:border-gray-600'
                          }`}
                        >
                          <div className="flex items-center space-x-3">
                            <div 
                              className="w-12 h-12 rounded-lg flex items-center justify-center"
                              style={{ backgroundColor: template.style.accentColor + '20' }}
                            >
                              <Palette className="w-6 h-6" style={{ color: template.style.accentColor }} />
                            </div>
                            <div>
                              <h3 className="font-semibold">{template.name}</h3>
                              <p className="text-sm text-gray-600 dark:text-gray-400">
                                {template.description}
                              </p>
                            </div>
                          </div>
                        </button>
                      ))}
                    </div>

                    <div className="mt-6">
                      <label className="flex items-center space-x-3">
                        <input
                          type="checkbox"
                          checked={currentRequest.includeImages || true}
                          onChange={(e) => updateRequestField('includeImages', e.target.checked)}
                          className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                        />
                        <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
                          Include relevant images from Pexels
                        </span>
                      </label>
                    </div>
                  </div>
                </div>

                <div className="mt-8 flex justify-center space-x-4">
                  <button
                    onClick={handlePreview}
                    disabled={!currentRequest.title || isLoadingPreview}
                    className="flex items-center space-x-2 px-6 py-3 bg-gray-500 text-white rounded-xl hover:bg-gray-600 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                  >
                    {isLoadingPreview ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      <Eye className="w-5 h-5" />
                    )}
                    <span>Preview Content</span>
                  </button>
                  
                  <button
                    onClick={handleGenerate}
                    disabled={!currentRequest.title || isGenerating}
                    className="flex items-center space-x-2 px-8 py-3 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl hover:from-blue-600 hover:to-purple-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-lg"
                  >
                    {isGenerating ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      <Zap className="w-5 h-5" />
                    )}
                    <span>Generate Presentation</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {step === 2 && (
            <motion.div
              key="step2"
              initial={{ opacity: 0, x: 50 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -50 }}
              className="max-w-6xl mx-auto"
            >
              <div className="glass rounded-2xl p-8 shadow-xl">
                <div className="text-center mb-8">
                  <h2 className="text-3xl font-bold mb-4">Preview Your Content</h2>
                  <p className="text-gray-600 dark:text-gray-400">
                    Review the generated content before creating your presentation
                  </p>
                </div>

                <div className="grid gap-6 mb-8">
                  {previewSlides.map((slide, index) => (
                    <motion.div
                      key={index}
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: index * 0.1 }}
                      className="bg-white dark:bg-gray-800 rounded-xl p-6 shadow-lg border border-gray-200 dark:border-gray-700"
                    >
                      <div className="flex items-start space-x-4">
                        <div className="flex-shrink-0 w-8 h-8 bg-blue-500 text-white rounded-lg flex items-center justify-center text-sm font-bold">
                          {index + 1}
                        </div>
                        <div className="flex-1">
                          <h3 className="text-xl font-bold mb-3 text-gray-900 dark:text-white">
                            {slide.header}
                          </h3>
                          <div className="text-gray-600 dark:text-gray-300 whitespace-pre-line">
                            {slide.content}
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  ))}
                </div>

                <div className="flex justify-center space-x-4">
                  <button
                    onClick={() => setStep(1)}
                    className="px-6 py-3 bg-gray-500 text-white rounded-xl hover:bg-gray-600 transition-all"
                  >
                    Back to Edit
                  </button>
                  
                  <button
                    onClick={handleGenerate}
                    disabled={isGenerating}
                    className="flex items-center space-x-2 px-8 py-3 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl hover:from-blue-600 hover:to-purple-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-lg"
                  >
                    {isGenerating ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      <FileText className="w-5 h-5" />
                    )}
                    <span>Create Presentation</span>
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {step === 3 && (
            <motion.div
              key="step3"
              initial={{ opacity: 0, y: 50 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -50 }}
              className="max-w-4xl mx-auto text-center"
            >
              <div className="glass rounded-2xl p-8 shadow-xl">
                <motion.div
                  initial={{ scale: 0 }}
                  animate={{ scale: 1 }}
                  transition={{ delay: 0.2, type: "spring" }}
                  className="w-20 h-20 bg-green-500 rounded-full flex items-center justify-center mx-auto mb-6"
                >
                  <CheckCircle className="w-10 h-10 text-white" />
                </motion.div>

                <h2 className="text-3xl font-bold mb-4">Presentation Ready!</h2>
                <p className="text-gray-600 dark:text-gray-400 mb-8">
                  Your AI-generated presentation has been created successfully. Download it now!
                </p>

                {generatedPresentation && (
                  <div className="bg-white dark:bg-gray-800 rounded-xl p-6 mb-8 text-left">
                    <h3 className="font-bold text-lg mb-2">{generatedPresentation.title}</h3>
                    <p className="text-sm text-gray-600 dark:text-gray-400 mb-4">
                      {generatedPresentation.slides?.length} slides • Generated with AI
                    </p>
                    
                    <div className="flex flex-wrap gap-2 mb-4">
                      <span className="px-3 py-1 bg-blue-100 dark:bg-blue-900 text-blue-800 dark:text-blue-200 rounded-full text-sm">
                        {currentRequest.template}
                      </span>
                      <span className="px-3 py-1 bg-purple-100 dark:bg-purple-900 text-purple-800 dark:text-purple-200 rounded-full text-sm">
                        {currentRequest.style}
                      </span>
                      <span className="px-3 py-1 bg-green-100 dark:bg-green-900 text-green-800 dark:text-green-200 rounded-full text-sm">
                        {currentRequest.tone}
                      </span>
                    </div>
                  </div>
                )}

                <div className="flex justify-center space-x-4">
                  <button
                    onClick={handleReset}
                    className="px-6 py-3 bg-gray-500 text-white rounded-xl hover:bg-gray-600 transition-all"
                  >
                    Create Another
                  </button>
                  
                  <button
                    onClick={handleDownload}
                    className="flex items-center space-x-2 px-8 py-3 bg-gradient-to-r from-green-500 to-blue-600 text-white rounded-xl hover:from-green-600 hover:to-blue-700 transition-all shadow-lg"
                  >
                    <Download className="w-5 h-5" />
                    <span>Download Presentation</span>
                  </button>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Generation Progress Overlay */}
        <AnimatePresence>
          {isGenerating && generationProgress && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50"
            >
              <motion.div
                initial={{ scale: 0.8, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.8, opacity: 0 }}
                className="bg-white dark:bg-gray-800 rounded-2xl p-8 max-w-md w-full mx-4 shadow-2xl"
              >
                <div className="text-center">
                  <div className="w-16 h-16 bg-gradient-to-r from-blue-500 to-purple-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <Loader2 className="w-8 h-8 text-white animate-spin" />
                  </div>
                  
                  <h3 className="text-xl font-bold mb-2">Creating Your Presentation</h3>
                  <p className="text-gray-600 dark:text-gray-400 mb-6">{generationProgress.message}</p>
                  
                  <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-2 mb-4">
                    <motion.div
                      className="bg-gradient-to-r from-blue-500 to-purple-600 h-2 rounded-full"
                      initial={{ width: 0 }}
                      animate={{ width: `${generationProgress.progress}%` }}
                      transition={{ duration: 0.5 }}
                    />
                  </div>
                  
                  <p className="text-sm text-gray-500">{generationProgress.progress}% complete</p>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Error Display */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: 50 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 50 }}
              className="fixed bottom-4 right-4 bg-red-500 text-white p-4 rounded-lg shadow-lg max-w-md"
            >
              <div className="flex items-start space-x-3">
                <AlertCircle className="w-5 h-5 mt-0.5 flex-shrink-0" />
                <div>
                  <h4 className="font-semibold">Error</h4>
                  <p className="text-sm">{error}</p>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
};

export default HomePage;
