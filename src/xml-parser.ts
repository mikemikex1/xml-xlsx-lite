/**
 * xml-xlsx-lite XML 解析工具
 */

/**
 * 簡單的 XML 解析器
 */
export class SimpleXMLParser {
  private text: string;
  private position: number;

  constructor(text: string) {
    this.text = text;
    this.position = 0;
  }

  /**
   * 解析 XML 文檔
   */
  parse(): XMLNode {
    return this.parseElement();
  }

  /**
   * 解析元素
   */
  private parseElement(): XMLNode {
    this.skipWhitespace();
    
    if (this.peek() !== '<') {
      throw new Error('Expected "<" at position ' + this.position);
    }
    
    this.consume(); // consume '<'
    
    // 檢查是否為結束標籤
    if (this.peek() === '/') {
      throw new Error('Unexpected closing tag at position ' + this.position);
    }
    
    // 讀取標籤名稱
    const tagName = this.readTagName();
    
    // 讀取屬性
    const attributes: Record<string, string> = {};
    this.skipWhitespace();
    
    while (this.peek() !== '>' && this.peek() !== '/' && !this.isEOF()) {
      const attr = this.readAttribute();
      attributes[attr.name] = attr.value;
      this.skipWhitespace();
    }
    
    // 檢查是否為自閉合標籤
    if (this.peek() === '/') {
      this.consume(); // consume '/'
      this.expect('>');
      return new XMLNode(tagName, attributes, [], '');
    }
    
    this.expect('>');
    
    // 讀取內容和子元素
    const children: XMLNode[] = [];
    let textContent = '';
    
    while (!this.isEOF()) {
      this.skipWhitespace();
      
      if (this.peek() === '<') {
        // 檢查是否為結束標籤
        if (this.peekString(2) === '</') {
          break;
        }
        // 解析子元素
        children.push(this.parseElement());
      } else {
        // 讀取文本內容
        textContent += this.readTextContent();
      }
    }
    
    // 讀取結束標籤
    this.expect('<');
    this.expect('/');
    const endTagName = this.readTagName();
    
    if (endTagName !== tagName) {
      throw new Error(`Mismatched tags: expected ${tagName}, got ${endTagName}`);
    }
    
    this.expect('>');
    
    return new XMLNode(tagName, attributes, children, textContent.trim());
  }

  /**
   * 讀取標籤名稱
   */
  private readTagName(): string {
    let name = '';
    while (!this.isEOF() && this.isNameChar(this.peek())) {
      name += this.consume();
    }
    return name;
  }

  /**
   * 讀取屬性
   */
  private readAttribute(): { name: string; value: string } {
    const name = this.readTagName();
    this.skipWhitespace();
    this.expect('=');
    this.skipWhitespace();
    
    const quote = this.peek();
    if (quote !== '"' && quote !== "'") {
      throw new Error('Expected quote at position ' + this.position);
    }
    
    this.consume(); // consume quote
    
    let value = '';
    while (!this.isEOF() && this.peek() !== quote) {
      value += this.consume();
    }
    
    this.expect(quote);
    
    return { name, value: this.unescapeXML(value) };
  }

  /**
   * 讀取文本內容
   */
  private readTextContent(): string {
    let content = '';
    while (!this.isEOF() && this.peek() !== '<') {
      content += this.consume();
    }
    return this.unescapeXML(content);
  }

  /**
   * 跳過空白字符
   */
  private skipWhitespace(): void {
    while (!this.isEOF() && this.isWhitespace(this.peek())) {
      this.consume();
    }
  }

  /**
   * 查看當前字符
   */
  private peek(): string {
    return this.position < this.text.length ? this.text[this.position] : '';
  }

  /**
   * 查看接下來的字符串
   */
  private peekString(length: number): string {
    return this.text.substring(this.position, this.position + length);
  }

  /**
   * 消費當前字符
   */
  private consume(): string {
    if (this.isEOF()) {
      throw new Error('Unexpected end of input');
    }
    return this.text[this.position++];
  }

  /**
   * 期望特定字符
   */
  private expect(char: string): void {
    if (this.peek() !== char) {
      throw new Error(`Expected "${char}" at position ${this.position}, got "${this.peek()}"`);
    }
    this.consume();
  }

  /**
   * 是否為文件結尾
   */
  private isEOF(): boolean {
    return this.position >= this.text.length;
  }

  /**
   * 是否為空白字符
   */
  private isWhitespace(char: string): boolean {
    return /\s/.test(char);
  }

  /**
   * 是否為名稱字符
   */
  private isNameChar(char: string): boolean {
    return /[a-zA-Z0-9:_-]/.test(char);
  }

  /**
   * XML 反轉義
   */
  private unescapeXML(text: string): string {
    return text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }
}

/**
 * XML 節點類別
 */
export class XMLNode {
  constructor(
    public tagName: string,
    public attributes: Record<string, string>,
    public children: XMLNode[],
    public textContent: string
  ) {}

  /**
   * 根據標籤名稱查找子元素
   */
  findChild(tagName: string): XMLNode | undefined {
    return this.children.find(child => child.tagName === tagName);
  }

  /**
   * 根據標籤名稱查找所有子元素
   */
  findChildren(tagName: string): XMLNode[] {
    return this.children.filter(child => child.tagName === tagName);
  }

  /**
   * 獲取屬性值
   */
  getAttribute(name: string): string | undefined {
    return this.attributes[name];
  }

  /**
   * 獲取文本內容
   */
  getText(): string {
    return this.textContent;
  }

  /**
   * 遞歸查找元素
   */
  findDeep(tagName: string): XMLNode | undefined {
    if (this.tagName === tagName) {
      return this;
    }
    
    for (const child of this.children) {
      const found = child.findDeep(tagName);
      if (found) {
        return found;
      }
    }
    
    return undefined;
  }

  /**
   * 遞歸查找所有元素
   */
  findAllDeep(tagName: string): XMLNode[] {
    const results: XMLNode[] = [];
    
    if (this.tagName === tagName) {
      results.push(this);
    }
    
    for (const child of this.children) {
      results.push(...child.findAllDeep(tagName));
    }
    
    return results;
  }
}

/**
 * 解析 XML 字符串
 */
export function parseXML(xmlText: string): XMLNode {
  const parser = new SimpleXMLParser(xmlText);
  return parser.parse();
}
