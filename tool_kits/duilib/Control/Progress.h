#ifndef UI_CONTROL_PROGRESS_H_
#define UI_CONTROL_PROGRESS_H_

#pragma once

namespace ui
{

class UILIB_API Progress : public LabelTemplate<Control>
{
public:
	Progress();

	/// ��д���෽�����ṩ���Ի����ܣ���ο���������
	virtual void SetAttribute(const std::wstring& strName, const std::wstring& strValue) override;
	virtual void PaintStatusImage(IRenderContext* pRender) override;
	virtual void ClearImageCache() override;

	/**
	 * @brief �ж��Ƿ���ˮƽ������
	 * @return ���� true ��ˮƽ��������false Ϊ��ֱ������
	 */
	bool IsHorizontal();

	/**
	 * @brief ����ˮƽ��ֱ������
	 * @param[in] bHorizontal Ϊ true ʱ����Ϊˮƽ��������false ʱ����Ϊ��ֱ��������Ĭ��Ϊ true
	 * @return ��
	 */
	void SetHorizontal(bool bHorizontal = true);

	/**
	 * @brief ��ȡ��������Сֵ
	 * @return ���ؽ�������Сֵ
	 */
	uint64_t GetMinValue() const;

	/**
	 * @brief ���ý�������Сֵ
	 * @param[in] nMin ��Сֵ��ֵ
	 * @return ��
	 */
	void SetMinValue(uint64_t nMin);

	/**
	 * @brief ��ȡ���������ֵ
	 * @return ���ؽ��������ֵ
	 */
	uint64_t GetMaxValue() const;

	/**
	 * @brief ���ý��������ֵ
	 * @param[in] nMax Ҫ���õ����ֵ
	 * @return ��
	 */
	void SetMaxValue(uint64_t nMax);

	/**
	 * @brief ��ȡ��ǰ���Ȱٷֱ�
	 * @return ���ص�ǰ���Ȱٷֱ�
	 */
	uint64_t GetValue() const;

	/**
	 * @brief ���õ�ǰ���Ȱٷֱ�
	 * @param[in] nValue Ҫ���õİٷֱ���ֵ
	 * @return ��
	 */
	void SetValue(uint64_t nValue);

	/**
	 * @brief ������ǰ��ͼƬ�Ƿ�������ʾ
	 * @return ���� true Ϊ������ʾ��false Ϊ��������ʾ
	 */
	bool IsStretchForeImage();

	/**
	 * @brief ���ý�����ǰ��ͼƬ�Ƿ�������ʾ
	 * @param[in] bStretchForeImage true Ϊ������ʾ��false Ϊ��������ʾ
	 * @return ��
	 */
	void SetStretchForeImage(bool bStretchForeImage = true);

	/**
	 * @brief ��ȡ����������ͼƬ
	 * @return ���ر���ͼƬλ��
	 */
	std::wstring GetProgressImage() const;

	/**
	 * @brief ���ý���������ͼƬ
	 * @param[in] strImage ͼƬ��ַ
	 * @return ��
	 */
	void SetProgressImage(const std::wstring& strImage);

	/**
	 * @brief ��ȡ������������ɫ
	 * @return ���ر�����ɫ���ַ���ֵ����Ӧ global.xml �е�ָ��ɫֵ
	 */
	std::wstring GetProgressColor() const;

	/**
	 * @brief ���ý�����������ɫ
	 * @param[in] Ҫ���õı�����ɫ�ַ��������ַ��������� global.xml �д���
	 * @return ��
	 */
	void SetProgressColor(const std::wstring& strProgressColor);

	/**
	 * @brief ��ȡ������λ��
	 * @return ���ؽ�������ǰλ��
	 */
	virtual UiRect GetProgressPos();

protected:
	bool m_bHorizontal;
	bool m_bStretchForeImage;
	uint64_t m_nMax;
	uint64_t m_nMin;
	uint64_t m_nValue;
	Image m_progressImage;
	std::wstring m_sProgressColor;
	std::wstring m_sProgressImageModify;
};

} // namespace ui

#endif // UI_CONTROL_PROGRESS_H_
