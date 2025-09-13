// 封面检测器 - 基于格式特征而非内容
class CoverDetector {
  constructor() {
    this.config = {
      // 格式特征检测阈值
      minWhitespaceElements: 3,     // 最少空白元素数
      minCenteredTexts: 2,          // 最少居中文本数
      maxTitleLength: 100,           // 标题最大长度
      leadingContentLength: 1500,   // 检查前部内容长度
      
      // 大字体检测
      largeFontSizes: {
        pt: 20,  // 20pt以上
        px: 30   // 30px以上
      }
    };
  }
  
  // 主检测方法
  detect(html) {
    console.log('开始检测封面格式特征...');
    
    // 先检查文档开头是否有典型封面特征
    const firstContent = html.substring(0, 2000);
    
    // 检查是否有多个居中的短文本（像您的截图中的封面）
    const centerPattern = /<p[^>]*>([^<]{1,100})<\/p>/gi;
    const matches = [...firstContent.matchAll(centerPattern)];
    
    if (matches.length >= 2) {
      const texts = matches.map(m => m[1].trim());
      console.log('找到居中文本:', texts.slice(0, 5));
      
      // 检查是否有典型的封面关键词
      const hasCoverKeywords = texts.some(t => 
        t.includes('API') || t.includes('报') || t.includes('审计') || 
        t.includes('公司') || t.includes('系统') || t.includes('专题')
      );
      
      // 检查是否有日期格式
      const hasDate = texts.some(t => /\d{4}年\d{1,2}月\d{1,2}日/.test(t));
      
      // 检查是否有安全室等部门名称
      const hasDepartment = texts.some(t => t.includes('安全') || t.includes('室'));
      
      if ((hasCoverKeywords && hasDate) || (texts.length >= 3 && hasDepartment)) {
        console.log('检测到封面特征!');
        return true;
      }
    }
    
    // 1. 检测空白行和间距
    if (this.hasSignificantWhitespace(html)) {
      console.log('检测到显著空白特征');
      return true;
    }
    
    // 2. 检测大字体标题
    if (this.hasLargeTitles(html)) {
      console.log('检测到大字体标题');
      return true;
    }
    
    // 3. 检测居中布局
    if (this.hasCenteredLayout(html)) {
      console.log('检测到居中布局');
      return true;
    }
    
    // 4. 检测稀疏内容（文字少但占用空间大）
    if (this.hasSparseContent(html)) {
      console.log('检测到稀疏内容布局');
      return true;
    }
    
    return false;
  }
  
  // 检测显著的空白
  hasSignificantWhitespace(html) {
    const leadingContent = html.substring(0, 500);
    
    // 统计各种空白元素
    const emptyParagraphs = (leadingContent.match(/<p[^>]*>\s*<\/p>/gi) || []).length;
    const lineBreaks = (leadingContent.match(/<br[^>]*>/gi) || []).length;
    const nbspGroups = (leadingContent.match(/(&nbsp;){3,}/g) || []).length;
    const emptyDivs = (leadingContent.match(/<div[^>]*>\s*<\/div>/gi) || []).length;
    
    const totalWhitespace = emptyParagraphs + lineBreaks + nbspGroups + emptyDivs;
    
    console.log('空白统计:', {
      emptyParagraphs,
      lineBreaks,
      nbspGroups,
      emptyDivs,
      total: totalWhitespace
    });
    
    return totalWhitespace >= this.config.minWhitespaceElements;
  }
  
  // 检测大字体标题
  hasLargeTitles(html) {
    const leadingContent = html.substring(0, 1000);
    
    // 检测h1标签（通常是大标题）
    if (/<h1[^>]*>/i.test(leadingContent)) {
      return true;
    }
    
    // 检测大字体样式
    const fontSizeMatch = leadingContent.match(/font-size:\s*(\d+)(pt|px)/gi);
    if (fontSizeMatch) {
      for (const match of fontSizeMatch) {
        const sizeMatch = match.match(/(\d+)(pt|px)/i);
        if (sizeMatch) {
          const size = parseInt(sizeMatch[1]);
          const unit = sizeMatch[2].toLowerCase();
          
          if (unit === 'pt' && size >= this.config.largeFontSizes.pt) {
            console.log(`发现大字体: ${size}pt`);
            return true;
          }
          if (unit === 'px' && size >= this.config.largeFontSizes.px) {
            console.log(`发现大字体: ${size}px`);
            return true;
          }
        }
      }
    }
    
    return false;
  }
  
  // 检测居中布局
  hasCenteredLayout(html) {
    const leadingContent = html.substring(0, this.config.leadingContentLength);
    
    // 统计居中元素
    const centerStyles = (leadingContent.match(/text-align:\s*center/gi) || []).length;
    const centerTags = (leadingContent.match(/<center[^>]*>/gi) || []).length;
    
    const totalCentered = centerStyles + centerTags;
    
    console.log('居中元素统计:', {
      centerStyles,
      centerTags,
      total: totalCentered
    });
    
    return totalCentered >= this.config.minCenteredTexts;
  }
  
  // 检测稀疏内容
  hasSparseContent(html) {
    const leadingContent = html.substring(0, this.config.leadingContentLength);
    
    // 移除HTML标签
    const textContent = leadingContent.replace(/<[^>]*>/g, '').trim();
    const textLength = textContent.length;
    const htmlLength = leadingContent.length;
    
    // 计算文本密度
    const textDensity = textLength / htmlLength;
    
    console.log('内容密度分析:', {
      textLength,
      htmlLength,
      density: textDensity.toFixed(2)
    });
    
    // 如果文本密度低于20%，可能是封面（很多格式，很少文字）
    return textDensity < 0.2 && textLength < 200;
  }
  
  // 提取封面内容
  extractCoverContent(html) {
    const content = {
      titles: [],
      date: null
    };
    
    const leadingContent = html.substring(0, this.config.leadingContentLength);
    
    // 提取h1-h3标题
    const h1Matches = leadingContent.match(/<h1[^>]*>([^<]+)<\/h1>/gi) || [];
    const h2Matches = leadingContent.match(/<h2[^>]*>([^<]+)<\/h2>/gi) || [];
    const h3Matches = leadingContent.match(/<h3[^>]*>([^<]+)<\/h3>/gi) || [];
    
    // 提取居中段落
    const centerMatches = leadingContent.match(/<p[^>]*style="[^"]*text-align:\s*center[^"]*"[^>]*>([^<]+)<\/p>/gi) || [];
    
    // 合并所有标题
    [...h1Matches, ...h2Matches, ...h3Matches, ...centerMatches].forEach(match => {
      const text = match.replace(/<[^>]*>/g, '').trim();
      if (text && text.length <= this.config.maxTitleLength) {
        content.titles.push(text);
      }
    });
    
    // 检测日期
    const datePattern = /[\(（]?\d{4}年\d{1,2}月\d{1,2}日[\)）]?/;
    const dateMatch = content.titles.find(t => datePattern.test(t));
    if (dateMatch) {
      content.date = dateMatch;
      content.titles = content.titles.filter(t => t !== dateMatch);
    }
    
    console.log('提取的封面内容:', content);
    return content;
  }
  
  // 查找第一个实质内容的位置
  findFirstRealContent(html) {
    // 查找表格
    const tablePos = html.indexOf('<table');
    
    // 查找较长的段落（超过100字符）
    const longParagraph = html.match(/<p[^>]*>[^<]{100,}<\/p>/i);
    const paragraphPos = longParagraph ? html.indexOf(longParagraph[0]) : -1;
    
    // 查找列表
    const ulPos = html.indexOf('<ul');
    const olPos = html.indexOf('<ol');
    const listPos = Math.min(
      ulPos >= 0 ? ulPos : Infinity,
      olPos >= 0 ? olPos : Infinity
    );
    
    // 返回最早的位置
    const positions = [tablePos, paragraphPos, listPos].filter(p => p > 0);
    const firstContent = positions.length > 0 ? Math.min(...positions) : -1;
    
    console.log('第一个实质内容位置:', firstContent);
    return firstContent;
  }
}

module.exports = CoverDetector;